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
        if not record.po_number:
            issues.append("Missing PO Number")
        
        if not record.description and not record.item_number:
            issues.append("Missing both Description and Item Number")
        
        if record.qty is None or record.qty <= 0:
            issues.append("Missing or invalid Quantity")
        
        # Check optional but important fields
        if not record.delivery_date:
            issues.append("Missing Delivery Date (optional)")
        
        if record.source_type == "pure" and not record.event_code:
            issues.append("PURE file missing Event Code (optional)")
        
        # Assign validation status and notes
        if not issues:
            record.validation_status = "valid"
            result.valid_records += 1
        elif len(issues) == 1 and "optional" in issues[0].lower():
            record.validation_status = "warning"
            record.validation_notes = issues
            result.warnings += 1
        else:
            # Errors outweigh warnings
            has_errors = any("optional" not in issue.lower() for issue in issues)
            if has_errors:
                record.validation_status = "error"
                result.errors += 1
            else:
                record.validation_status = "warning"
                result.warnings += 1
            record.validation_notes = issues
        
        validated.append(record)
        result.details.append({
            "index": i,
            "po": record.po_number,
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
