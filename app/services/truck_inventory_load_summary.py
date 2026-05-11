"""Summaries for item-based truck loads."""

from __future__ import annotations

from app.models.truck_inventory_record import TruckInventoryRecord


def get_load_summary(records: list[TruckInventoryRecord]) -> dict:
    """Summarize normalized rows by KKG Load #."""
    summary: dict[str, dict] = {}
    for record in records:
        load_number = record.kkg_load_number or "UNASSIGNED"
        if load_number not in summary:
            summary[load_number] = {
                "kkg_load_number": load_number,
                "rows": 0,
                "item_numbers": set(),
                "total_qty": 0,
                "total_weight": 0.0,
                "validation_status": "valid",
                "validation_notes": [],
            }

        qty = record.qty or 0
        summary[load_number]["rows"] += 1
        if record.item_number:
            summary[load_number]["item_numbers"].add(record.item_number)
        summary[load_number]["total_qty"] += qty
        summary[load_number]["total_weight"] += (record.item_weight or 0.0) * qty

        if record.validation_status == "error":
            summary[load_number]["validation_status"] = "error"
        elif (
            record.validation_status == "warning"
            and summary[load_number]["validation_status"] != "error"
        ):
            summary[load_number]["validation_status"] = "warning"

        for note in record.validation_notes:
            if note not in summary[load_number]["validation_notes"]:
                summary[load_number]["validation_notes"].append(note)

    for data in summary.values():
        data["item_count"] = len(data.pop("item_numbers"))
        data["validation_notes"] = "; ".join(data["validation_notes"])

    return summary
