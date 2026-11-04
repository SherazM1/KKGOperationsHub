"""Validate declared pallet groups without treating totals as item quantities."""

import math
from app.services.bol_excel_headers import compact_header


def number(value):
    try:
        parsed = float(str(value).replace(",", "").strip())
        return parsed if math.isfinite(parsed) and parsed >= 0 else None
    except (ValueError, TypeError):
        return None


def reconcile_pallet_groups(rows, *, pallet_field="plt_qty", weight_field="total_weight"):
    """A populated Total PLT QTY begins a new group inside the shipment."""
    blocks, block, previous_key = [], [], None
    for row in rows:
        # A truck/load group can include several destination BOLs. Its total
        # spans the contiguous rows until the next declaration or load.
        key = (row.kk_load,) if row.kk_load.strip() else (row.bol_number, row.dc_number)
        if block and (key != previous_key or row.group_pallet_total.strip() or row.group_weight_total.strip()):
            blocks.append(block)
            block = []
        block.append(row)
        previous_key = key
    if block:
        blocks.append(block)
    for block in blocks:
        first = block[0]
        if not first.group_pallet_total.strip():
            continue
        declared = number(first.group_pallet_total)
        if len(block) == 1 and not getattr(first, pallet_field).strip() and declared is not None:
            setattr(first, pallet_field, first.group_pallet_total)
            first.column_mapping[pallet_field] = first.column_mapping.get("group_pallet_total", "Total PLT QTY")
            first.parsing_notes = [note for note in first.parsing_notes
                                   if note != "No separate pallet/skid count supplied; review required."]
            first.parsing_notes.append("Single-item group's pallet count comes from its declared group total.")
        counts = [number(getattr(row, pallet_field)) for row in block]
        if declared is None or any(value is None for value in counts):
            message = "Group pallet total cannot be reconciled with its item pallet counts; review required."
        elif not math.isclose(sum(counts), declared, rel_tol=1e-9, abs_tol=1e-6):
            message = f"Group pallet total is {declared:g}, but item pallet counts sum to {sum(counts):g}; review required."
        else:
            message = ""
            first.parsing_notes.append(
                f"Pallet group at Excel row {first.source_row_number}: item counts sum to {declared:g}; group total counted once."
            )
        if message:
            for row in block:
                row.parsing_notes.append(message)
            continue
        # Some GC sheets supply one load weight at the top of the group, with
        # blanks below it. Only derive line weights when the source explicitly
        # supplies weight PER PALLET and the group reconciles exactly.
        weights = [number(getattr(row, weight_field)) for row in block]
        if len(block) < 2 or sum(value is not None for value in weights) != 1:
            continue
        if not all(compact_header(row.column_mapping.get("weight_each", "")) in
                   {"PLTWEIGHT", "PALLETWEIGHT", "WEIGHTPERPALLET"} for row in block):
            continue
        each = [number(row.weight_each) for row in block]
        if any(value is None for value in each):
            continue
        calculated = [count * unit_weight for count, unit_weight in zip(counts, each)]
        supplied = next(value for value in weights if value is not None)
        if not math.isclose(sum(calculated), supplied, rel_tol=1e-9, abs_tol=0.01):
            for row in block:
                row.parsing_notes.append("Group weight differs from item pallets times pallet weight; review required.")
            continue
        first.group_weight_total = f"{supplied:g}"
        for row, weight in zip(block, calculated):
            setattr(row, weight_field, f"{weight:g}")
            source_weight = row.column_mapping.get(weight_field, "Weight")
            row.column_mapping["group_weight_total"] = source_weight
            row.column_mapping[weight_field] = "Derived from item pallet count times PLT Weight"
            row.parsing_notes.append("Item weight derived from PLT QTY times PLT Weight; source group weight reconciled once.")
    for block in blocks:
        if not block[0].group_weight_total.strip():
            continue
        declared = number(block[0].group_weight_total)
        weights = [number(getattr(row, weight_field)) for row in block]
        if declared is None or any(value is None for value in weights) or not math.isclose(
                sum(weights), declared, rel_tol=1e-9, abs_tol=0.01):
            for row in block:
                row.parsing_notes.append("Declared load weight does not reconcile with the item weights; review required.")
        else:
            block[0].parsing_notes.append(f"Declared load weight {declared:g} reconciled once against item weights.")
    return rows
