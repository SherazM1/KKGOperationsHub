"""Normalize parsed inventory data to standard schema."""

from __future__ import annotations

from app.models.truck_inventory_record import TruckInventoryRecord


def normalize_pure_row(row: dict, source_file: str) -> TruckInventoryRecord:
    """
    Normalize a PURE file row to standard schema.
    
    PURE files typically have:
    - PO Number
    - Event Code
    - Delivery Date
    - DC fields
    - Item fields
    - Quantity
    """
    record = TruckInventoryRecord(
        source_type="pure",
        source_file=source_file,
        po_number=_get_field(row, ["PO", "PO_Number", "PO Number", "po_number"], ""),
        event_code=_get_field(row, ["Event Code", "event_code", "Event"], ""),
        delivery_date=_get_field(row, ["Delivery Date", "delivery_date", "Delivery"], None),
        dc_number=_get_field(row, ["DC", "DC_Number", "DC Number", "dc_number"], None),
        dc_name=_get_field(row, ["DC Name", "dc_name"], None),
        receiver_name=_get_field(row, ["Receiver", "receiver_name", "Receiver Name"], None),
        ship_to_address_1=_get_field(row, ["Address 1", "address_1", "Address"], None),
        ship_to_address_2=_get_field(row, ["Address 2", "address_2"], None),
        ship_to_city=_get_field(row, ["City", "city"], None),
        ship_to_state=_get_field(row, ["State", "state"], None),
        ship_to_zip=_get_field(row, ["Zip", "zip", "ZIP"], None),
        item_number=_get_field(row, ["Item", "item_number", "Item Number"], None),
        upc=_get_field(row, ["UPC", "upc"], None),
        description=_get_field(row, ["Description", "description", "Desc"], None),
        qty=_parse_int(row, ["Qty", "Quantity", "qty", "quantity"]),
        uom=_get_field(row, ["UOM", "uom", "Unit"], "EA"),
        unit_weight=_parse_float(row, ["Unit Weight", "unit_weight", "Weight/Unit"]),
        total_weight=_parse_float(row, ["Total Weight", "total_weight"]),
    )
    
    # Set initial load group from event code or default
    record.load_group = record.event_code or "GENERAL"
    
    return record


def normalize_cdw_row(row: dict, source_file: str) -> TruckInventoryRecord:
    """
    Normalize a CDW file row to standard schema.
    
    CDW files typically have:
    - Purchase Order fields
    - Line item details
    - Quantities
    - Descriptions
    """
    record = TruckInventoryRecord(
        source_type="cdw",
        source_file=source_file,
        po_number=_get_field(row, ["PO", "Purchase Order", "PO_Number", "po"], ""),
        po_date=_get_field(row, ["PO Date", "po_date", "Order Date"], None),
        delivery_date=_get_field(row, ["Delivery Date", "delivery_date", "Due Date"], None),
        dc_number=_get_field(row, ["DC", "Destination", "dc_number"], None),
        receiver_name=_get_field(row, ["Receiver", "Ship To", "receiver_name"], None),
        ship_to_city=_get_field(row, ["City", "city"], None),
        ship_to_state=_get_field(row, ["State", "state"], None),
        ship_to_zip=_get_field(row, ["Zip", "zip", "ZIP"], None),
        item_number=_get_field(row, ["Item", "Item Number", "SKU", "sku"], None),
        upc=_get_field(row, ["UPC", "upc", "EAN"], None),
        description=_get_field(row, ["Description", "description", "Product", "product_name"], None),
        qty=_parse_int(row, ["Qty", "Quantity", "qty", "quantity", "Qty Ordered"]),
        uom=_get_field(row, ["UOM", "uom", "Unit"], "EA"),
        unit_weight=_parse_float(row, ["Unit Weight", "unit_weight", "Weight/Unit"]),
    )
    
    # Calculate total weight if available
    if record.unit_weight and record.qty:
        record.total_weight = record.unit_weight * record.qty
    
    # Default load group
    record.load_group = "GENERAL"
    
    return record


def normalize_combined_row(row: dict, source_file: str) -> TruckInventoryRecord:
    """Normalize a combined load sheet row."""
    # Combined sheets may already be partially normalized
    record = TruckInventoryRecord(
        source_type="combined",
        source_file=source_file,
        po_number=_get_field(row, ["PO", "PO_Number", "po"], ""),
        delivery_date=_get_field(row, ["Delivery Date", "delivery_date"], None),
        dc_number=_get_field(row, ["DC", "dc_number"], None),
        dc_name=_get_field(row, ["DC Name", "dc_name"], None),
        item_number=_get_field(row, ["Item", "item_number"], None),
        description=_get_field(row, ["Description", "description"], None),
        qty=_parse_int(row, ["Qty", "Quantity", "qty"]),
        estimated_pallets=_parse_float(row, ["Pallets", "pallets", "Estimated Pallets"]),
        load_group=_get_field(row, ["Load Group", "load_group", "Group"], "GENERAL"),
    )
    
    return record


def normalize_rows(
    rows: list[dict],
    source_file: str,
    file_type: str,
) -> list[TruckInventoryRecord]:
    """
    Normalize a list of rows based on detected or specified file type.
    
    Args:
        rows: List of dictionaries from parsed Excel
        source_file: Original filename
        file_type: "pure", "cdw", or "combined_load_sheet"
        
    Returns:
        List of normalized TruckInventoryRecord objects
    """
    normalized = []
    
    for row in rows:
        # Skip empty rows
        if not row or all(v == "" or v is None for v in row.values()):
            continue
        
        if file_type == "pure":
            record = normalize_pure_row(row, source_file)
        elif file_type == "cdw":
            record = normalize_cdw_row(row, source_file)
        elif file_type == "combined_load_sheet":
            record = normalize_combined_row(row, source_file)
        else:
            # Fallback: try to detect from row content
            if _looks_like_pure(row):
                record = normalize_pure_row(row, source_file)
            elif _looks_like_cdw(row):
                record = normalize_cdw_row(row, source_file)
            else:
                record = normalize_combined_row(row, source_file)
        
        normalized.append(record)
    
    return normalized


def _get_field(row: dict, possible_names: list[str], default=None) -> str | None:
    """Get a field value from a row, checking multiple possible column names."""
    for name in possible_names:
        if name in row and row[name] and row[name] != "":
            return str(row[name]).strip()
    return default


def _parse_int(row: dict, possible_names: list[str]) -> int | None:
    """Parse an integer field from a row."""
    for name in possible_names:
        if name in row:
            try:
                return int(float(row[name]))
            except (ValueError, TypeError):
                continue
    return None


def _parse_float(row: dict, possible_names: list[str]) -> float | None:
    """Parse a float field from a row."""
    for name in possible_names:
        if name in row:
            try:
                return float(row[name])
            except (ValueError, TypeError):
                continue
    return None


def _looks_like_pure(row: dict) -> bool:
    """Check if a row looks like PURE data."""
    keys_lower = [str(k).lower() for k in row.keys()]
    pure_indicators = ["event", "po", "dc", "delivery"]
    return sum(1 for ind in pure_indicators if any(ind in k for k in keys_lower)) >= 2


def _looks_like_cdw(row: dict) -> bool:
    """Check if a row looks like CDW data."""
    keys_lower = [str(k).lower() for k in row.keys()]
    cdw_indicators = ["purchase", "order", "sku", "customer"]
    return sum(1 for ind in cdw_indicators if any(ind in k for k in keys_lower)) >= 1
