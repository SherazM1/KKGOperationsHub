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


def export_load_summary_csv(summary: dict) -> bytes:
    """Export item-based load summary to CSV."""
    rows = []
    for load_number, data in sorted(summary.items()):
        rows.append({
            "KKG Load #": load_number,
            "Rows": data["rows"],
            "Unique Item # Count": data["item_count"],
            "Total Qty": data["total_qty"],
            "Total Weight": data["total_weight"],
            "Validation Status": data["validation_status"],
            "Validation Notes": data["validation_notes"],
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
            "KKG Load #": truck.kkg_load_number,
            "Preset": truck.truck_preset,
            "Truck Length": truck.truck_length,
            "Truck Width": truck.truck_width,
            "Truck Height": truck.truck_height,
            "Truck Max Weight": truck.truck_max_weight,
            "Used Floor Area": truck.total_used_pallets,
            "Floor Area Capacity": truck.total_capacity_pallets,
            "Floor Utilization %": truck.utilization_percent,
            "Total Weight": truck.total_weight,
            "Weight Utilization %": truck.weight_utilization_percent,
            "Item Qty": truck.items_count,
            "Spatial Fit": truck.spatial_fit,
            "Weight Fit": truck.weight_fit,
            "Validation Status": truck.validation_status,
            "Validation Notes": ";".join(truck.validation_notes),
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
                "KKG Load #": box.kkg_load_number,
                "Retailer PO #": box.retailer_po_number,
                "Item #": box.item_number,
                "Qty": box.row_qty,
                "Unit Index": box.unit_index,
                "Description": box.description,
                "Color": box.color,
                "X Position": box.x,
                "Y Position": box.y,
                "Width": box.width,
                "Height": box.height,
                "Item Length": box.item_length,
                "Item Width": box.item_width,
                "Item Height": box.item_height,
                "Item Weight": box.item_weight,
                "Stack Level": box.stack_level,
                "Stack Qty": box.stack_qty,
            })
    
    df = pd.DataFrame(rows)
    
    csv_buffer = io.BytesIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    return csv_buffer.getvalue()


def export_required_columns_csv(records: list[TruckInventoryRecord]) -> bytes:
    """Export the minimum required operations columns as CSV."""
    rows = [
        {
            "KKG Load #": record.kkg_load_number,
            "Retailer PO #": record.retailer_po_number or record.po_number,
            "Item #": record.item_number,
            "Qty": record.qty,
        }
        for record in records
    ]
    df = pd.DataFrame(rows)

    csv_buffer = io.BytesIO()
    df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    return csv_buffer.getvalue()


def export_required_columns_excel(records: list[TruckInventoryRecord]) -> bytes:
    """Export only KKG Load #, Retailer PO #, Item #, and Qty as an Excel file."""
    rows = [
        {
            "KKG Load #": record.kkg_load_number,
            "Retailer PO #": record.retailer_po_number or record.po_number,
            "Item #": record.item_number,
            "Qty": record.qty,
        }
        for record in records
    ]
    df = pd.DataFrame(rows)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Truck Export")
    buffer.seek(0)
    return buffer.getvalue()


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
