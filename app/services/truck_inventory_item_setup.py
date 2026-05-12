"""Item setup helpers for Truck Inventory."""

from __future__ import annotations

from copy import deepcopy

from app.models.truck_inventory_record import TruckInventoryRecord
from app.utils.truck_presets import (
    ITEM_COLOR_BY_ITEM_NUMBER,
    ITEM_PRESETS,
    ITEM_TYPE_BY_ITEM_NUMBER,
    UNKNOWN_ITEM_COLOR,
)


def get_distinct_item_numbers(records: list[TruckInventoryRecord]) -> list[str]:
    """Return sorted distinct Item # values from normalized business rows."""
    return sorted({record.item_number for record in records if record.item_number})


def build_default_item_setup(records: list[TruckInventoryRecord]) -> list[dict]:
    """Create editable item setup rows keyed by Item #."""
    setup_rows = []
    for item_number in get_distinct_item_numbers(records):
        preset_key = infer_item_preset_key(item_number)
        preset = ITEM_PRESETS[preset_key]
        setup_rows.append({
            "Item #": item_number,
            "Preset": preset.name,
            "Length": preset.length,
            "Width": preset.width,
            "Height": preset.height,
            "Weight": preset.weight,
            "Is Stackable?": "No",
            "Stack Qty": 1,
            "Color": infer_item_color(item_number),
        })
    return setup_rows


def merge_item_setup(existing_rows: list[dict], records: list[TruckInventoryRecord]) -> list[dict]:
    """Keep user-entered item setup while adding rows for newly parsed items."""
    existing_by_item = {str(row.get("Item #", "")): row for row in existing_rows if row.get("Item #")}
    merged = []
    defaults_by_item = {row["Item #"]: row for row in build_default_item_setup(records)}

    for item_number in get_distinct_item_numbers(records):
        if item_number in existing_by_item:
            row = deepcopy(existing_by_item[item_number])
            row["Item #"] = item_number
            merged.append(row)
        else:
            merged.append(defaults_by_item[item_number])
    return merged


def apply_item_setup(
    records: list[TruckInventoryRecord],
    setup_rows: list[dict],
) -> tuple[list[TruckInventoryRecord], list[str]]:
    """Apply UI-provided physical item setup values to normalized records."""
    issues = validate_item_setup(records, setup_rows)
    setup_by_item = {str(row.get("Item #", "")): row for row in setup_rows if row.get("Item #")}

    for record in records:
        item_number = record.item_number or ""
        setup = setup_by_item.get(item_number)
        if not setup:
            continue

        record.item_length = _to_float(setup.get("Length"))
        record.item_width = _to_float(setup.get("Width"))
        record.item_height = _to_float(setup.get("Height"))
        record.item_weight = _to_float(setup.get("Weight"))
        record.is_stackable = str(setup.get("Is Stackable?", "No")).lower() == "yes"
        record.stack_qty = _effective_stack_qty(setup.get("Stack Qty"), record.is_stackable)
        record.color_group = str(setup.get("Color") or record.item_number or "UNKNOWN")
        record.total_weight = (record.item_weight or 0.0) * (record.qty or 0)

    return records, sorted(issues)


def validate_item_setup(records: list[TruckInventoryRecord], setup_rows: list[dict]) -> list[str]:
    """Validate editable item setup rows without mutating records."""
    setup_by_item = {str(row.get("Item #", "")): row for row in setup_rows if row.get("Item #")}
    issues: set[str] = set()

    for item_number in get_distinct_item_numbers(records):
        setup = setup_by_item.get(item_number)
        if not setup:
            issues.add(f"Item # {item_number}: missing item setup")
            continue

        missing_fields = []
        for label, key in [
            ("length", "Length"),
            ("width", "Width"),
            ("height", "Height"),
            ("weight", "Weight"),
        ]:
            if not _positive(_to_float(setup.get(key))):
                missing_fields.append(label)
        if missing_fields:
            issues.add(f"Item # {item_number}: missing or invalid {', '.join(missing_fields)}")

        is_stackable = str(setup.get("Is Stackable?", "No")).lower() == "yes"
        if is_stackable and _parse_stack_qty(setup.get("Stack Qty")) is None:
            issues.add(f"Item # {item_number}: stack qty must be a whole number greater than 0")

    return sorted(issues)


def preset_to_setup_values(preset_name: str) -> dict:
    """Return editable setup values for a selected item preset name."""
    preset_key = preset_name.strip().lower()
    if preset_key == "custom":
        return {}
    for preset in ITEM_PRESETS.values():
        if preset.name.lower() == preset_key:
            return {
                "Length": preset.length,
                "Width": preset.width,
                "Height": preset.height,
                "Weight": preset.weight,
                "Is Stackable?": "Yes" if preset.is_stackable else "No",
                "Stack Qty": preset.stack_qty,
            }
    return {}


def infer_item_preset_key(item_number: str) -> str:
    """Infer a known item preset from the current sample item mappings."""
    return ITEM_TYPE_BY_ITEM_NUMBER.get(str(item_number).strip(), "custom")


def infer_item_color(item_number: str) -> str:
    """Infer a starter item color from current sample mappings."""
    return ITEM_COLOR_BY_ITEM_NUMBER.get(str(item_number).strip(), UNKNOWN_ITEM_COLOR)


def _effective_stack_qty(value, is_stackable: bool) -> int:
    if not is_stackable:
        return 1
    parsed = _parse_stack_qty(value)
    return parsed or 1


def _parse_stack_qty(value) -> int | None:
    try:
        stack_qty = int(float(value))
    except (TypeError, ValueError):
        return None
    return stack_qty if stack_qty > 0 else None


def _to_float(value) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _positive(value: float | None) -> bool:
    return value is not None and value > 0
