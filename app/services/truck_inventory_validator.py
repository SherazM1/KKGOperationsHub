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
