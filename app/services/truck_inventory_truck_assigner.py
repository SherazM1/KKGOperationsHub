"""Assign loads to trucks sequentially."""

from __future__ import annotations

from app.models.truck_inventory_record import TruckInventoryRecord
from app.models.truck_summary import BoxLayout, TruckSummary
from app.utils.truck_presets import get_preset, get_color_for_group


def assign_to_trucks(
    records: list[TruckInventoryRecord],
    preset_key: str = "53ft_dry_van",
    grouping_rule: str = "load_group",
) -> list[TruckSummary]:
    """
    Assign inventory records to trucks sequentially.
    
    Args:
        records: List of normalized and pallet-calculated records
        preset_key: Truck preset key (from truck_presets.TRUCK_PRESETS)
        grouping_rule: How to group records ("load_group", "po_number", "none")
        
    Returns:
        List of TruckSummary objects with assignments
    """
    preset = get_preset(preset_key)
    trucks: list[TruckSummary] = []
    
    # Group records if needed
    grouped = _group_records(records, grouping_rule)
    
    # Track which truck we're filling
    current_truck: TruckSummary | None = None
    truck_id_counter = 1
    
    # Color mapping for load groups
    group_color_map: dict[str, int] = {}
    
    for group_key, group_records in grouped.items():
        # Calculate total pallets needed for this group
        total_pallets = sum(r.estimated_pallets or 0.0 for r in group_records)
        
        # See if we can fit this group in the current truck
        if current_truck is None or \
           current_truck.total_used_pallets + total_pallets > current_truck.total_capacity_pallets:
            # Create a new truck
            current_truck = TruckSummary(
                truck_id=truck_id_counter,
                truck_preset=preset.name,
                total_capacity_pallets=preset.pallet_capacity,
                total_used_pallets=0.0,
                utilization_percent=0.0,
                remaining_capacity=preset.pallet_capacity,
            )
            trucks.append(current_truck)
            truck_id_counter += 1
        
        # Add this group to the current truck
        box_id = len(current_truck.boxes) + 1
        load_group = group_key
        pallets = total_pallets
        
        # Get color for this load group
        color = get_color_for_group(load_group, group_color_map)
        
        # Create box layout (simple layout: stack vertically)
        box = BoxLayout(
            box_id=box_id,
            load_group=load_group,
            pallet_count=pallets,
            description=f"{load_group} ({len(group_records)} items)",
            color=color,
            x=0.1,
            y=0.1 + (box_id - 1) * 0.12,  # Stack boxes vertically
            width=0.8,
            height=0.1,
        )
        
        current_truck.boxes.append(box)
        current_truck.total_used_pallets += pallets
        current_truck.remaining_capacity = current_truck.total_capacity_pallets - current_truck.total_used_pallets
        
        if current_truck.total_capacity_pallets > 0:
            current_truck.utilization_percent = (
                current_truck.total_used_pallets / current_truck.total_capacity_pallets * 100
            )
        
        if load_group not in current_truck.load_groups:
            current_truck.load_groups.append(load_group)
        
        current_truck.items_count += len(group_records)
    
    return trucks


def _group_records(
    records: list[TruckInventoryRecord],
    grouping_rule: str,
) -> dict[str, list[TruckInventoryRecord]]:
    """
    Group records based on grouping rule.
    
    Args:
        records: List of records
        grouping_rule: "load_group", "po_number", "none", or custom field name
        
    Returns:
        Dictionary mapping group key to list of records
    """
    grouped: dict[str, list[TruckInventoryRecord]] = {}
    
    for record in records:
        if grouping_rule == "load_group":
            key = record.load_group or "GENERAL"
        elif grouping_rule == "po_number":
            key = record.po_number or "UNSPECIFIED"
        elif grouping_rule == "none":
            # Each record is its own group
            key = f"item_{record.item_number}_{len(grouped)}"
        else:
            # Default to load_group
            key = record.load_group or "GENERAL"
        
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(record)
    
    return grouped


def get_truck_summary_stats(trucks: list[TruckSummary]) -> dict:
    """Get overall statistics about truck assignments."""
    total_capacity = sum(t.total_capacity_pallets for t in trucks)
    total_used = sum(t.total_used_pallets for t in trucks)
    total_remaining = sum(t.remaining_capacity for t in trucks)
    
    avg_utilization = 0.0
    if trucks:
        avg_utilization = sum(t.utilization_percent for t in trucks) / len(trucks)
    
    return {
        "total_trucks": len(trucks),
        "total_capacity": total_capacity,
        "total_used_pallets": total_used,
        "total_remaining": total_remaining,
        "average_utilization": avg_utilization,
    }
