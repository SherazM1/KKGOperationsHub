"""Assign KKG loads to trucks using simple item-based placement."""

from __future__ import annotations

from collections import OrderedDict

from app.models.truck_inventory_record import TruckInventoryRecord
from app.models.truck_summary import BoxLayout, TruckSummary
from app.utils.truck_presets import ITEM_PRESETS, ITEM_TYPE_BY_ITEM_NUMBER, get_color_for_item, get_preset


def assign_to_trucks(
    records: list[TruckInventoryRecord],
    preset_key: str = "dry_van",
    grouping_rule: str = "kkg_load_number",
    operational_weight_threshold_lbs: float | None = None,
) -> list[TruckSummary]:
    """
    Build one truck plan per KKG Load #.

    MVP assumptions:
    - one visual box is one grouped stack unit
    - stacking is same-item only, based on each item's stack qty
    - item rows and KKG loads are not split across trucks
    - item dimensions are interpreted in inches
    """
    preset = get_preset(preset_key)
    trucks: list[TruckSummary] = []
    item_color_map: dict[str, int] = {}

    for truck_id, (load_number, group_records) in enumerate(
        _group_records(records, grouping_rule).items(),
        start=1,
    ):
        _clear_previous_fit_notes(group_records)
        truck_length = _first_positive(group_records, "truck_length") or preset.length_in
        truck_width = _first_positive(group_records, "truck_width") or preset.width_in
        truck_height = _first_positive(group_records, "truck_height") or preset.height_in
        truck_max_weight = (
            operational_weight_threshold_lbs
            or _first_positive(group_records, "truck_max_weight")
            or preset.operational_weight_threshold_lbs
        )
        total_weight = sum((r.item_weight or 0.0) * (r.qty or 0) for r in group_records)
        preflight_notes = _validate_load_records(
            group_records,
            truck_height=truck_height,
            truck_max_weight=truck_max_weight,
            total_weight=total_weight,
        )

        boxes, spatial_fit, spatial_notes = _place_load_items(
            group_records,
            truck_length,
            truck_width,
            truck_height,
            item_color_map,
        )

        weight_fit = total_weight <= truck_max_weight
        notes = _unique_notes([*preflight_notes, *spatial_notes])
        if not weight_fit:
            notes = _unique_notes([
                *notes,
                f"Weight: load is {total_weight:,.1f} lbs, threshold is {truck_max_weight:,.1f} lbs",
            ])

        preflight_fit = not preflight_notes
        status = "valid" if preflight_fit and spatial_fit and weight_fit else "error"
        floor_area = truck_length * truck_width
        used_floor_area = _used_floor_area(group_records)
        utilization = (used_floor_area / floor_area * 100.0) if floor_area > 0 else 0.0
        remaining_area = max(floor_area - used_floor_area, 0.0)

        truck = TruckSummary(
            truck_id=truck_id,
            truck_preset=_first_text(group_records, "truck_type") or preset.name,
            total_capacity_pallets=floor_area,
            total_used_pallets=used_floor_area,
            utilization_percent=utilization,
            remaining_capacity=remaining_area,
            load_groups=[load_number],
            items_count=sum(r.qty or 0 for r in group_records),
            boxes=boxes,
            notes="; ".join(notes),
            kkg_load_number=load_number,
            truck_length=truck_length,
            truck_width=truck_width,
            truck_height=truck_height,
            truck_max_weight=truck_max_weight,
            total_weight=total_weight,
            weight_utilization_percent=(total_weight / truck_max_weight * 100.0) if truck_max_weight > 0 else 0.0,
            spatial_fit=spatial_fit and preflight_fit,
            weight_fit=weight_fit,
            validation_status=status,
            validation_notes=notes,
        )
        trucks.append(truck)

        for record in group_records:
            record.validation_notes = _merge_notes(record.validation_notes, notes)
            if status == "error":
                record.validation_status = "error"

    return trucks


def _place_load_items(
    records: list[TruckInventoryRecord],
    truck_length: float,
    truck_width: float,
    truck_height: float,
    item_color_map: dict[str, int],
) -> tuple[list[BoxLayout], bool, list[str]]:
    boxes: list[BoxLayout] = []
    notes: list[str] = []
    x = 0.0
    y = 0.0
    row_depth = 0.0
    box_id = 1

    for record in records:
        qty = record.qty or 0
        length = record.item_length or 0.0
        width = record.item_width or 0.0
        height = record.item_height or 0.0
        item_number = record.item_number or "UNKNOWN"
        effective_stack_qty = max(1, record.stack_qty if record.is_stackable else 1)
        stack_height = height * effective_stack_qty

        row_fits = True
        if length <= 0 or width <= 0 or height <= 0:
            row_fits = False
        if height > truck_height:
            row_fits = False
        if record.is_stackable and stack_height > truck_height:
            row_fits = False
        if length > truck_length or width > truck_width:
            _add_note(notes, f"Spatial: Item # {item_number} footprint {length:g} x {width:g} exceeds truck floor")
            row_fits = False

        row_start_box_count = len(boxes)
        floor_slots = _floor_slots_for_qty(qty, effective_stack_qty)
        unit_index = 1
        for slot_index in range(1, floor_slots + 1):
            if x + length > truck_length:
                x = 0.0
                y += row_depth
                row_depth = 0.0

            if y + width > truck_width:
                _add_note(notes, f"Spatial: Item # {item_number} row qty {qty} does not fit without splitting")
                row_fits = False
                break

            represented_qty = min(effective_stack_qty, qty - unit_index + 1)
            grouped_height = height * represented_qty
            color = _item_color(record, item_color_map)
            item_type = _item_type(item_number)
            boxes.append(
                BoxLayout(
                    box_id=box_id,
                    load_group=record.kkg_load_number or record.load_group or "UNASSIGNED",
                    pallet_count=float(represented_qty),
                    description=(
                        f"Item {item_number} group {slot_index} of {floor_slots}; "
                        f"represents {represented_qty} unit(s)"
                    ),
                    color=color,
                    x=x / truck_length if truck_length else 0.0,
                    y=y / truck_width if truck_width else 0.0,
                    width=length / truck_length if truck_length else 0.0,
                    height=width / truck_width if truck_width else 0.0,
                    kkg_load_number=record.kkg_load_number or "",
                    retailer_po_number=record.retailer_po_number or record.po_number,
                    item_number=item_number,
                    row_qty=qty,
                    represented_qty=represented_qty,
                    unit_index=slot_index,
                    item_length=length,
                    item_width=width,
                    item_height=height,
                    grouped_height=grouped_height,
                    item_weight=record.item_weight or 0.0,
                    stack_level=represented_qty,
                    stack_qty=effective_stack_qty,
                    item_type=item_type,
                )
            )
            box_id += 1
            unit_index += represented_qty
            x += length
            row_depth = max(row_depth, width)

        if not row_fits:
            # A row cannot be split, so remove any partial units from that row.
            del boxes[row_start_box_count:]

    return boxes, not notes, _unique_notes(notes)


def _validate_load_records(
    records: list[TruckInventoryRecord],
    truck_height: float,
    truck_max_weight: float,
    total_weight: float,
) -> list[str]:
    notes: list[str] = []
    item_issues: OrderedDict[str, set[str]] = OrderedDict()

    for record in records:
        item_number = record.item_number or "UNKNOWN"
        issues = item_issues.setdefault(item_number, set())
        if not _positive(record.item_length):
            issues.add("length")
        if not _positive(record.item_width):
            issues.add("width")
        if not _positive(record.item_height):
            issues.add("height")
        if not _positive(record.item_weight):
            issues.add("weight")

        stack_qty = record.stack_qty if record.is_stackable else 1
        if record.is_stackable and (stack_qty is None or stack_qty <= 0):
            _add_note(notes, f"Stacking: Item # {item_number} stack qty must be greater than 0")
        if _positive(record.item_height) and (record.item_height or 0.0) > truck_height:
            _add_note(
                notes,
                f"Dimensions: Item # {item_number} height {record.item_height:g} exceeds truck height {truck_height:g}",
            )
        if record.is_stackable and _positive(record.item_height) and stack_qty > 0:
            stack_height = (record.item_height or 0.0) * stack_qty
            if stack_height > truck_height:
                max_stack = max(1, int(truck_height // (record.item_height or 1.0)))
                _add_note(
                    notes,
                    f"Stacking: Item # {item_number} stack height {stack_height:g} exceeds truck height "
                    f"{truck_height:g}; max stack qty is {max_stack}",
                )

    for item_number, missing_fields in item_issues.items():
        if missing_fields:
            _add_note(
                notes,
                f"Item setup: Item # {item_number} missing or invalid {', '.join(sorted(missing_fields))}",
            )

    if total_weight > truck_max_weight:
        _add_note(notes, f"Weight: load is {total_weight:,.1f} lbs, threshold is {truck_max_weight:,.1f} lbs")

    return _unique_notes(notes)


def _floor_slots_for_qty(qty: int, stack_qty: int) -> int:
    if qty <= 0:
        return 0
    return (qty + stack_qty - 1) // stack_qty


def _used_floor_area(records: list[TruckInventoryRecord]) -> float:
    total = 0.0
    for record in records:
        stack_qty = max(1, record.stack_qty if record.is_stackable else 1)
        total += (record.item_length or 0.0) * (record.item_width or 0.0) * _floor_slots_for_qty(record.qty or 0, stack_qty)
    return total


def _positive(value: float | None) -> bool:
    return value is not None and value > 0


def _add_note(notes: list[str], note: str) -> None:
    if note not in notes:
        notes.append(note)


def _unique_notes(notes: list[str]) -> list[str]:
    unique = []
    for note in notes:
        if note and note not in unique:
            unique.append(note)
    return unique


def _item_color(record: TruckInventoryRecord, item_color_map: dict[str, int]) -> str:
    if record.color_group and str(record.color_group).startswith("#"):
        return str(record.color_group)
    return get_color_for_item(record.item_number or "UNKNOWN", item_color_map)


def _item_type(item_number: str) -> str:
    preset_key = ITEM_TYPE_BY_ITEM_NUMBER.get(str(item_number).strip(), "")
    if not preset_key:
        return ""
    return ITEM_PRESETS[preset_key].name


def _group_records(
    records: list[TruckInventoryRecord],
    grouping_rule: str,
) -> OrderedDict[str, list[TruckInventoryRecord]]:
    grouped: OrderedDict[str, list[TruckInventoryRecord]] = OrderedDict()
    for record in records:
        if grouping_rule == "po_number":
            key = record.retailer_po_number or record.po_number or "UNSPECIFIED"
        elif grouping_rule == "none":
            key = f"{record.kkg_load_number or 'UNASSIGNED'}-{record.item_number or 'UNKNOWN'}"
        else:
            key = record.kkg_load_number or record.load_group or "UNASSIGNED"
        grouped.setdefault(key, []).append(record)
    return grouped


def _first_positive(records: list[TruckInventoryRecord], attr: str) -> float | None:
    for record in records:
        value = getattr(record, attr)
        if value is not None and value > 0:
            return value
    return None


def _first_text(records: list[TruckInventoryRecord], attr: str) -> str | None:
    for record in records:
        value = getattr(record, attr)
        if value:
            return value
    return None


def _merge_notes(existing: list[str], new_notes: list[str]) -> list[str]:
    merged = list(existing)
    for note in new_notes:
        if note not in merged:
            merged.append(note)
    return merged


def _clear_previous_fit_notes(records: list[TruckInventoryRecord]) -> None:
    fit_prefixes = ("Item setup:", "Dimensions:", "Stacking:", "Weight:", "Spatial:")
    for record in records:
        record.validation_notes = [
            note for note in record.validation_notes
            if not str(note).startswith(fit_prefixes)
        ]
        if record.validation_status == "error" and not record.validation_notes:
            record.validation_status = "valid"


def get_truck_summary_stats(trucks: list[TruckSummary]) -> dict:
    """Get overall statistics about item-based truck assignments."""
    total_capacity = sum(t.total_capacity_pallets for t in trucks)
    total_used = sum(t.total_used_pallets for t in trucks)
    total_remaining = sum(t.remaining_capacity for t in trucks)
    total_weight = sum(t.total_weight for t in trucks)

    avg_utilization = 0.0
    if trucks:
        avg_utilization = sum(t.utilization_percent for t in trucks) / len(trucks)

    return {
        "total_trucks": len(trucks),
        "total_capacity": total_capacity,
        "total_used_pallets": total_used,
        "total_remaining": total_remaining,
        "total_weight": total_weight,
        "average_utilization": avg_utilization,
        "failed_loads": sum(1 for t in trucks if t.validation_status == "error"),
    }
