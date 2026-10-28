"""Item setup helpers for Truck Inventory."""

from __future__ import annotations

from copy import deepcopy
from math import isfinite
from app.services.truck_inventory_stacking import is_half_pallet

from app.models.truck_inventory_record import TruckInventoryRecord
from app.utils.truck_presets import (
    ITEM_COLOR_BY_ITEM_NUMBER,
    ITEM_PRESETS,
    ITEM_SETUP_COLOR_OPTIONS,
    ITEM_TYPE_BY_ITEM_NUMBER,
    UNKNOWN_ITEM_COLOR,
)


LEGACY_PRESET_NAMES = {
    "pure": "Full Pallet",
    "cdw": "Half Pallet",
}


def normalize_editor_rows(rows: list[dict]) -> list[dict]:
    """Translate checkbox values into the saved stacking configuration."""
    normalized_rows = []
    for row in rows:
        normalized = dict(row)

        if "Double Stack" in normalized:
            double_stack = bool(normalized.pop("Double Stack"))
            normalized["Is Stackable?"] = "Yes" if double_stack else "No"
            normalized["Stack Qty"] = 2 if double_stack else 1
            if not double_stack:
                normalized["Stack By Dimension"] = False
        if str(normalized.get("Color", "")) not in ITEM_SETUP_COLOR_OPTIONS:
            normalized["Color"] = "Orange"
        normalized_rows.append(normalized)
    return normalized_rows


def get_distinct_item_numbers(records: list[TruckInventoryRecord]) -> list[str]:
    """Return sorted distinct Item # values from normalized business rows."""
    return sorted({record.item_number for record in records if record.item_number})


def get_dimension_groups(setup_rows: list[dict]) -> dict[tuple[float, float, float], list[str]]:
    """Group items by their exact Item Setup length, width and height in inches."""
    groups = {}
    for row in setup_rows:
        dimensions = tuple(_to_float(row.get(field)) for field in ("Length", "Width", "Height"))
        if row.get("Item #") and all(_positive(value) and isfinite(value) for value in dimensions):
            groups.setdefault(dimensions, []).append(str(row["Item #"]))
    return dict(sorted(groups.items()))


def apply_dimension_group_stacking(
    setup_rows: list[dict], records: list[TruckInventoryRecord],
    dimensions: list[tuple[float, float, float]], enabled: bool,
) -> tuple[list[dict], list[str]]:
    """Update stacking for selected groups without changing physical dimensions."""
    selected = set(dimensions)
    updated = deepcopy(setup_rows)
    for row in updated:
        size = tuple(_to_float(row.get(field)) for field in ("Length", "Width", "Height"))
        if size not in selected:
            continue
        row["Is Stackable?"] = "Yes" if enabled else "No"
        row["Stack Qty"] = 2 if enabled else 1
        row["Stack By Dimension"] = bool(enabled)
    return updated, []


def build_default_item_setup(records: list[TruckInventoryRecord]) -> list[dict]:
    """Create editable item setup rows keyed by Item #."""
    setup_rows = []
    for item_number in get_distinct_item_numbers(records):
        preset_key = infer_item_preset_key(item_number)
        preset = ITEM_PRESETS[preset_key]
        item_records = [record for record in records if record.item_number == item_number]
        imported = {}
        for label, attr in [("Length", "item_length"), ("Width", "item_width"), ("Height", "item_height"), ("Weight", "item_weight")]:
            def input_value(record):
                original = getattr(record, 'unrotated_' + attr, None)
                return original if original is not None else getattr(record, attr)
            values = {input_value(record) for record in item_records if input_value(record) is not None}
            if values:
                imported[label] = next(iter(values)) if len(values) == 1 else 0.0
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
            **imported,
        })
    return setup_rows


def merge_item_setup(existing_rows: list[dict], records: list[TruckInventoryRecord]) -> list[dict]:
    """Keep user-entered item setup while adding rows for newly parsed items."""
    existing_by_item = {
        str(row.get("Item #", "")): _normalize_setup_row(row)
        for row in existing_rows
        if row.get("Item #")
    }
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
        record.unrotated_item_length = record.item_length
        record.unrotated_item_width = record.item_width
        record.item_height = _to_float(setup.get("Height"))
        record.item_weight = _to_float(setup.get("Weight"))
        record.is_stackable = str(setup.get("Is Stackable?", "No")).lower() == "yes"
        record.stack_by_dimension = bool(setup.get("Stack By Dimension", False)) and record.is_stackable
        record.stack_qty = _effective_stack_qty(setup.get("Stack Qty"), record.is_stackable)
        if is_half_pallet(record, str(setup.get("Preset", ""))):
            record.pallet_type = "Half Pallet"

            # Rotate once from the saved input footprint, not from a previously
            # rotated plan. This also rotates imported 48 x 24 half pallets.
            record.item_length, record.item_width = (
                record.item_width,
                record.item_length,
            )
        record.color_group = setup_color_to_hex(setup.get("Color"))
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
    preset_label = _normalize_preset_name(preset_name)
    if preset_label.lower() == "custom":
        return {}
    for preset in ITEM_PRESETS.values():
        if preset.name.lower() == preset_label.lower():
            return {
                "Length": preset.length,
                "Width": preset.width,
            }
    return {}


def infer_item_preset_key(item_number: str) -> str:
    """Infer a known item preset from the current sample item mappings."""
    return ITEM_TYPE_BY_ITEM_NUMBER.get(str(item_number).strip(), "custom")


def infer_item_color(item_number: str) -> str:
    """Infer a starter item color from current sample mappings."""
    return ITEM_COLOR_BY_ITEM_NUMBER.get(str(item_number).strip(), UNKNOWN_ITEM_COLOR)


def setup_color_to_hex(color_name: object) -> str:
    """Resolve the manual Item Setup color label to the render color."""
    return ITEM_SETUP_COLOR_OPTIONS.get(str(color_name or "").strip(), ITEM_SETUP_COLOR_OPTIONS[UNKNOWN_ITEM_COLOR])


def _effective_stack_qty(value, is_stackable: bool) -> int:
    if not is_stackable:
        return 1
    parsed = _parse_stack_qty(value)
    return parsed or 1


def _parse_stack_qty(value) -> int | None:
    try:
        number = float(value)
        if not isfinite(number) or not number.is_integer():
            return None
        stack_qty = int(number)
    except (TypeError, ValueError, OverflowError):
        return None
    return stack_qty if stack_qty > 0 else None


def _to_float(value) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _positive(value: float | None) -> bool:
    return value is not None and value > 0


def _normalize_setup_row(row: dict) -> dict:
    normalized = deepcopy(row)
    normalized["Preset"] = _normalize_preset_name(str(normalized.get("Preset", "Custom")))
    normalized["Color"] = _normalize_color_name(normalized.get("Color"))
    return normalized


def _normalize_preset_name(preset_name: str) -> str:
    preset_key = str(preset_name or "").strip().lower()
    return LEGACY_PRESET_NAMES.get(preset_key, preset_name or "Custom")


def _normalize_color_name(color_name: object) -> str:
    value = str(color_name or "").strip()
    if value in ITEM_SETUP_COLOR_OPTIONS:
        return value
    for name, hex_value in ITEM_SETUP_COLOR_OPTIONS.items():
        if value.lower() == hex_value.lower():
            return name
    return UNKNOWN_ITEM_COLOR
