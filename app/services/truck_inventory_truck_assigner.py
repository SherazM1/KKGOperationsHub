"""Assign KKG loads to trucks using simple item-based placement."""

from __future__ import annotations

from collections import OrderedDict

from app.models.truck_inventory_record import TruckInventoryRecord
from app.models.truck_summary import BoxLayout, TruckSummary
from app.utils.truck_presets import get_color_for_item, get_preset


def assign_to_trucks(
    records: list[TruckInventoryRecord],
    preset_key: str = "53ft_dry_van",
    grouping_rule: str = "kkg_load_number",
) -> list[TruckSummary]:
    """
    Build one truck plan per KKG Load #.

    MVP assumptions:
    - one visual box is one item unit
    - stacking is one high
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
        truck_length = _first_positive(group_records, "truck_length") or preset.length_in
        truck_width = _first_positive(group_records, "truck_width") or preset.width_in
        truck_height = _first_positive(group_records, "truck_height") or preset.height_in
        truck_max_weight = _first_positive(group_records, "truck_max_weight") or preset.max_weight_lbs
        total_weight = sum((r.item_weight or 0.0) * (r.qty or 0) for r in group_records)

        boxes, spatial_fit, spatial_notes = _place_load_items(
            group_records,
            truck_length,
            truck_width,
            truck_height,
            item_color_map,
        )

        weight_fit = total_weight <= truck_max_weight
        notes = list(spatial_notes)
        if not weight_fit:
            notes.append(f"Load weight {total_weight:.1f} exceeds max {truck_max_weight:.1f}")

        status = "valid" if spatial_fit and weight_fit else "error"
        floor_area = truck_length * truck_width
        used_floor_area = sum(
            (r.item_length or 0.0) * (r.item_width or 0.0) * (r.qty or 0)
            for r in group_records
        )
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
            spatial_fit=spatial_fit,
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

        row_fits = True
        if length <= 0 or width <= 0 or height <= 0:
            notes.append(f"Item {item_number} has missing dimensions")
            row_fits = False
        if height > truck_height:
            notes.append(f"Item {item_number} height {height:.1f} exceeds truck height {truck_height:.1f}")
            row_fits = False
        if length > truck_length or width > truck_width:
            notes.append(f"Item {item_number} footprint does not fit truck floor")
            row_fits = False

        row_start_box_count = len(boxes)
        for unit_index in range(1, qty + 1):
            if x + length > truck_length:
                x = 0.0
                y += row_depth
                row_depth = 0.0

            if y + width > truck_width:
                notes.append(f"Item row {item_number} qty {qty} does not fit without splitting")
                row_fits = False
                break

            color = get_color_for_item(item_number, item_color_map)
            boxes.append(
                BoxLayout(
                    box_id=box_id,
                    load_group=record.kkg_load_number or record.load_group or "UNASSIGNED",
                    pallet_count=1.0,
                    description=f"Item {item_number} unit {unit_index} of {qty}",
                    color=color,
                    x=x / truck_length if truck_length else 0.0,
                    y=y / truck_width if truck_width else 0.0,
                    width=length / truck_length if truck_length else 0.0,
                    height=width / truck_width if truck_width else 0.0,
                    kkg_load_number=record.kkg_load_number or "",
                    retailer_po_number=record.retailer_po_number or record.po_number,
                    item_number=item_number,
                    row_qty=qty,
                    unit_index=unit_index,
                    item_length=length,
                    item_width=width,
                    item_height=height,
                    item_weight=record.item_weight or 0.0,
                )
            )
            box_id += 1
            x += length
            row_depth = max(row_depth, width)

        if not row_fits:
            # A row cannot be split, so remove any partial units from that row.
            del boxes[row_start_box_count:]

    return boxes, not notes, notes


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
