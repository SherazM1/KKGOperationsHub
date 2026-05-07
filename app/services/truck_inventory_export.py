"""Export truck inventory data to CSV formats."""

from __future__ import annotations

import io
import pandas as pd

from app.models.truck_inventory_record import TruckInventoryRecord
from app.models.truck_summary import TruckSummary


def export_normalized_data_csv(records: list[TruckInventoryRecord]) -> bytes:
    """
    Export normalized records to CSV.
    
    Args:
        records: List of TruckInventoryRecord objects
        
    Returns:
        CSV file as bytes
    """
    data = [record.to_dict() for record in records]
    df = pd.DataFrame(data)
    
    csv_buffer = io.BytesIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    return csv_buffer.getvalue()


def export_pallet_summary_csv(summary: dict) -> bytes:
    """
    Export pallet summary to CSV.
    
    Args:
        summary: Dictionary mapping load_group to pallet counts
        
    Returns:
        CSV file as bytes
    """
    rows = []
    for group, data in sorted(summary.items()):
        rows.append({
            "Load Group": group,
            "Item Count": data["items"],
            "Total Pallets": data["total_pallets"],
            "Total Quantity": data["total_qty"],
        })
    
    df = pd.DataFrame(rows)
    
    csv_buffer = io.BytesIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    return csv_buffer.getvalue()


def export_truck_assignments_csv(trucks: list[TruckSummary]) -> bytes:
    """
    Export truck assignments and utilization to CSV.
    
    Args:
        trucks: List of TruckSummary objects
        
    Returns:
        CSV file as bytes
    """
    rows = []
    for truck in trucks:
        rows.append({
            "Truck ID": truck.truck_id,
            "Preset": truck.truck_preset,
            "Capacity": truck.total_capacity_pallets,
            "Used Pallets": truck.total_used_pallets,
            "Utilization %": truck.utilization_percent,
            "Remaining": truck.remaining_capacity,
            "Load Groups": ";".join(truck.load_groups),
            "Item Count": truck.items_count,
        })
    
    df = pd.DataFrame(rows)
    
    csv_buffer = io.BytesIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    return csv_buffer.getvalue()


def export_truck_boxes_csv(trucks: list[TruckSummary]) -> bytes:
    """
    Export detailed box/load placements for each truck.
    
    Args:
        trucks: List of TruckSummary objects
        
    Returns:
        CSV file as bytes
    """
    rows = []
    for truck in trucks:
        for box in truck.boxes:
            rows.append({
                "Truck ID": truck.truck_id,
                "Box ID": box.box_id,
                "Load Group": box.load_group,
                "Pallets": box.pallet_count,
                "Description": box.description,
                "Color": box.color,
                "X Position": box.x,
                "Y Position": box.y,
                "Width": box.width,
                "Height": box.height,
            })
    
    df = pd.DataFrame(rows)
    
    csv_buffer = io.BytesIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    return csv_buffer.getvalue()


def export_combined_report_csv(
    records: list[TruckInventoryRecord],
    trucks: list[TruckSummary],
) -> bytes:
    """
    Export a combined report with both normalized data and truck assignments.
    
    Args:
        records: List of TruckInventoryRecord objects
        trucks: List of TruckSummary objects
        
    Returns:
        CSV file as bytes
    """
    # Create a DataFrame with records, then add truck assignment info
    data = [record.to_dict() for record in records]
    df = pd.DataFrame(data)
    
    # Add truck assignment if available
    # This is a simplified version - in a real implementation,
    # you'd need to track which record goes to which truck
    df["truck_id"] = ""
    
    csv_buffer = io.BytesIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    return csv_buffer.getvalue()
