"""Shared pallet stacking restrictions."""
from app.models.truck_inventory_record import TruckInventoryRecord
from app.utils.truck_presets import ITEM_TYPE_BY_ITEM_NUMBER, TRUCK_PRESETS

def get_roof_clearance(preset_key: str, truck_name: str | None = None) -> float:
    preset = TRUCK_PRESETS.get(preset_key)
    if preset is None:
        preset = next((value for value in TRUCK_PRESETS.values() if value.name == truck_name), None)
    return preset.roof_clearance_in if preset else 3.0


def is_half_pallet(record: TruckInventoryRecord, preset: str = "") -> bool:
    labels = [record.pallet_type or "", preset,
              ITEM_TYPE_BY_ITEM_NUMBER.get(record.item_number or "", "")]
    return any("half" in label.lower().replace("_", " ").replace("-", " ")
               or "1/2" in label or "½" in label for label in labels)
