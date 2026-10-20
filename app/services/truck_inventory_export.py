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


def export_required_columns_excel(records: list[TruckInventoryRecord], trucks: list[TruckSummary] | None = None) -> bytes:
    """Export planning identifiers, retaining store inventory details when present."""
    def assignment_key(item, po):
        return (item.kkg_load_number or "", po or "", item.item_number or "",
                item.store_number or "", item.dc_number or "",
                item.item_length or 0, item.item_width or 0,
                item.item_height or 0, item.item_weight or 0)

    allocations = {}
    for truck in trucks or []:
        for box in truck.boxes:
            quantities = allocations.setdefault(assignment_key(box, box.retailer_po_number), {})
            quantities[truck.truck_id] = quantities.get(truck.truck_id, 0) + box.represented_qty
    
    rows = [
        {
            "KKG Load #": record.kkg_load_number,
            "Retailer PO #": record.retailer_po_number or record.po_number,
            "Item #": record.item_number,
            "Qty": record.qty,
        }
        for record in records
    ]
    if any(record.store_number is not None for record in records):
        for row, record in zip(rows, records):
            row.update({
                "Store #": record.store_number,
                "DC #": record.dc_number,
                "Pallet Type": record.pallet_type,
                "Pallet Length": record.item_length,
                "Pallet Width": record.item_width,
                "Pallet Height": record.item_height,
                "Weight": record.item_weight,
            })
    assigned_rows = []
    for row, record in zip(rows, records):
        remaining = record.qty or 0
        quantities = allocations.get(
            assignment_key(record, record.retailer_po_number or record.po_number), {}
        )
        for truck_id, available in quantities.items():
            qty = min(remaining, available)
            if qty <= 0:
                continue
            assigned_rows.append({**row, "KKG Load #": f"{record.kkg_load_number or ''}-{truck_id}", "Qty": qty})
            quantities[truck_id] -= qty
            remaining -= qty
        # Preserve unassigned quantities without implying a truck assignment.
        if remaining or not record.qty:
            assigned_rows.append({**row, "Qty": remaining if record.qty else record.qty})
    df = pd.DataFrame(assigned_rows)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Truck Export")
        if trucks:
            assignments = [{
                "Truck #": truck.truck_id, "Truck Type": truck.truck_preset,
                "KKG Load #": box.kkg_load_number, "Retailer PO #": box.retailer_po_number,
                "Store #": box.store_number, "DC #": box.dc_number,
                "Item #": box.item_number, "Qty": box.represented_qty,
                "Pallet Length": box.item_length, "Pallet Width": box.item_width,
                "Pallet Height": box.item_height,
            } for truck in trucks for box in truck.boxes]
            pd.DataFrame(assignments).to_excel(writer, index=False, sheet_name="Truck Assignments")
            issues = [{"Truck #": truck.truck_id, "Loads": truck.kkg_load_number,
                       "Qty Requiring Review": truck.items_count,
                       "Issues": "; ".join(truck.validation_notes)}
                      for truck in trucks if truck.validation_status == "error"]
            if issues:
                pd.DataFrame(issues).to_excel(writer, index=False, sheet_name="Unassigned Review")
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
