"""Shared pallet stacking restrictions."""
from app.models.truck_inventory_record import TruckInventoryRecord
from app.utils.truck_presets import ITEM_TYPE_BY_ITEM_NUMBER

STACK_ROOF_CLEARANCE_IN = 6.0


def is_half_pallet(record: TruckInventoryRecord, preset: str = "") -> bool:
    labels = [record.pallet_type or "", preset,
              ITEM_TYPE_BY_ITEM_NUMBER.get(record.item_number or "", "")]
    return any("half" in label.lower().replace("_", " ").replace("-", " ")
               or "1/2" in label or "½" in label for label in labels)
