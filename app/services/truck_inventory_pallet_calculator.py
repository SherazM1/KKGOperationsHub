"""Calculate estimated pallets from inventory records."""

from __future__ import annotations

from app.models.truck_inventory_record import TruckInventoryRecord


# Placeholder pallet calculation modes
PALLET_CALC_MODES = {
    "direct_qty": "Use quantity directly as pallet count",
    "qty_divided_by_cases": "Divide quantity by cases-per-pallet",
    "description_inference": "Infer from description (heuristic)",
    "manual_override": "Manual override / placeholder",
}

# Default cases per pallet for 'qty_divided_by_cases' mode
DEFAULT_CASES_PER_PALLET = 40


def estimate_pallets_direct_qty(record: TruckInventoryRecord) -> float:
    """Use quantity directly as pallet count (placeholder mode)."""
    if record.qty is None:
        return 0.0
    return float(record.qty)


def estimate_pallets_divided_by_cases(
    record: TruckInventoryRecord,
    cases_per_pallet: int = DEFAULT_CASES_PER_PALLET,
) -> float:
    """Divide quantity by cases-per-pallet."""
    if record.qty is None:
        return 0.0
    pallets = record.qty / cases_per_pallet
    return round(pallets, 2)


def estimate_pallets_from_description(record: TruckInventoryRecord) -> float:
    """
    Heuristic: infer pallet count from description keywords.
    
    Placeholder logic - this is where business rules will be added.
    """
    if not record.description:
        return 1.0  # Default to 1 pallet
    
    desc_lower = record.description.lower()
    
    # Heuristic: check for common pallet/case keywords
    if any(word in desc_lower for word in ["case", "cases", "pack", "packs"]):
        # Assume quantity is in cases
        return estimate_pallets_divided_by_cases(record)
    
    if any(word in desc_lower for word in ["pallet", "pallets"]):
        # Assume quantity is already in pallets
        if record.qty:
            return float(record.qty)
    
    # Default
    return 1.0


def estimate_pallets_manual_override(record: TruckInventoryRecord) -> float:
    """
    Manual override mode: use existing estimated_pallets if set, else calculate.
    
    This allows users to manually set pallet counts in the input.
    """
    if record.estimated_pallets is not None and record.estimated_pallets > 0:
        return record.estimated_pallets
    
    # Fall back to default calculation
    return estimate_pallets_direct_qty(record)


def calculate_pallets(
    records: list[TruckInventoryRecord],
    mode: str = "direct_qty",
    cases_per_pallet: int = DEFAULT_CASES_PER_PALLET,
) -> list[TruckInventoryRecord]:
    """
    Calculate estimated pallets for all records using specified mode.
    
    Args:
        records: List of TruckInventoryRecord objects
        mode: One of the PALLET_CALC_MODES keys
        cases_per_pallet: Number of cases per pallet (for divided mode)
        
    Returns:
        Records with estimated_pallets populated
    """
    for record in records:
        if mode == "direct_qty":
            record.estimated_pallets = estimate_pallets_direct_qty(record)
        elif mode == "qty_divided_by_cases":
            record.estimated_pallets = estimate_pallets_divided_by_cases(record, cases_per_pallet)
        elif mode == "description_inference":
            record.estimated_pallets = estimate_pallets_from_description(record)
        elif mode == "manual_override":
            record.estimated_pallets = estimate_pallets_manual_override(record)
        else:
            # Default
            record.estimated_pallets = estimate_pallets_direct_qty(record)
    
    return records


def get_pallet_summary(records: list[TruckInventoryRecord]) -> dict:
    """Get a summary of pallet counts by load group."""
    summary = {}
    
    for record in records:
        if record.load_group not in summary:
            summary[record.load_group] = {
                "group": record.load_group,
                "items": 0,
                "total_pallets": 0.0,
                "total_qty": 0,
            }
        
        summary[record.load_group]["items"] += 1
        summary[record.load_group]["total_pallets"] += record.estimated_pallets or 0.0
        summary[record.load_group]["total_qty"] += record.qty or 0
    
    return summary
