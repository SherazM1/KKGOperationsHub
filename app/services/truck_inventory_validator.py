"""Validate normalized truck inventory records."""

from __future__ import annotations

from app.models.truck_inventory_record import TruckInventoryRecord


class ValidationResult:
    """Result of validating records."""
    
    def __init__(self):
        self.total_records = 0
        self.valid_records = 0
        self.warnings = 0
        self.errors = 0
        self.details = []
    
    def to_dict(self):
        """Convert to dict for display."""
        return {
            "Total Records": self.total_records,
            "Valid": self.valid_records,
            "Warnings": self.warnings,
            "Errors": self.errors,
        }


def validate_records(records: list[TruckInventoryRecord]) -> tuple[list[TruckInventoryRecord], ValidationResult]:
    """
    Validate a list of normalized records.
    
    Args:
        records: List of TruckInventoryRecord objects
        
    Returns:
        Tuple of (validated records, validation result summary)
    """
    result = ValidationResult()
    result.total_records = len(records)
    
    validated = []
    
    for i, record in enumerate(records):
        issues = []
        
        # Check required fields
        if not record.kkg_load_number:
            issues.append("Missing KKG Load #")
        
        if not record.retailer_po_number and not record.po_number:
            issues.append("Missing Retailer PO #")
        
        if not record.item_number:
            issues.append("Missing Item #")
        
        if record.qty is None or record.qty <= 0:
            issues.append("Missing or invalid Quantity")

        for label, value in [
            ("length", record.item_length),
            ("width", record.item_width),
            ("height", record.item_height),
            ("weight", record.item_weight),
        ]:
            if value is None or value <= 0:
                issues.append(f"Missing or invalid item {label}")
        
        if record.truck_length is not None and record.truck_length <= 0:
            issues.append("Invalid truck length")
        if record.truck_width is not None and record.truck_width <= 0:
            issues.append("Invalid truck width")
        if record.truck_height is not None and record.truck_height <= 0:
            issues.append("Invalid truck height")
        if record.truck_max_weight is not None and record.truck_max_weight <= 0:
            issues.append("Invalid truck max weight")
        
        # Assign validation status and notes
        if not issues:
            record.validation_status = "valid"
            result.valid_records += 1
        else:
            record.validation_status = "error"
            result.errors += 1
            record.validation_notes = issues
        
        validated.append(record)
        result.details.append({
            "index": i,
            "kkg_load_number": record.kkg_load_number,
            "po": record.retailer_po_number or record.po_number,
            "item": record.item_number,
            "status": record.validation_status,
            "notes": "; ".join(record.validation_notes),
        })
    
    return validated, result


def get_validation_summary(result: ValidationResult) -> str:
    """Get a human-readable validation summary."""
    lines = [
        f"Total Records: {result.total_records}",
        f"Valid: {result.valid_records}",
        f"Warnings: {result.warnings}",
        f"Errors: {result.errors}",
    ]
    return " | ".join(lines)
