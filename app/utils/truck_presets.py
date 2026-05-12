"""Truck presets and configuration for inventory assignments."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class TruckPreset:
    """Represents a standard truck capacity preset."""

    name: str
    pallet_capacity: float  # Legacy display value retained for compatibility.
    length_in: float
    width_in: float
    height_in: float
    max_weight_lbs: float
    operational_weight_threshold_lbs: float
    description: str

    @property
    def length_ft(self) -> float:
        return self.length_in / 12.0

    @property
    def width_ft(self) -> float:
        return self.width_in / 12.0

    @property
    def height_ft(self) -> float:
        return self.height_in / 12.0


# Standard truck presets
TRUCK_PRESETS = {
    "dry_van": TruckPreset(
        name="Dry Van",
        pallet_capacity=26.0,
        length_in=636.0,
        width_in=96.0,
        height_in=102.0,
        max_weight_lbs=45000.0,
        operational_weight_threshold_lbs=45000.0,
        description="Dry van trailer",
    ),
    "reefer": TruckPreset(
        name="Reefer",
        pallet_capacity=22.0,
        length_in=515.75,
        width_in=90.0,
        height_in=96.0,
        max_weight_lbs=43000.0,
        operational_weight_threshold_lbs=43000.0,
        description="Refrigerated trailer",
    ),
    "box_truck_26ft": TruckPreset(
        name="Standard 26 Foot Box Truck",
        pallet_capacity=13.0,
        length_in=312.0,
        width_in=96.0,
        height_in=84.0,
        max_weight_lbs=8700.0,
        operational_weight_threshold_lbs=8700.0,
        description="Standard 26 foot box truck",
    ),
}


@dataclass(frozen=True)
class ItemPreset:
    """Physical defaults for common pallet footprints."""

    name: str
    length: float
    width: float
    height: float
    weight: float
    is_stackable: bool
    stack_qty: int


ITEM_PRESETS = {
    "full_pallet": ItemPreset(
        name="Full Pallet",
        length=40.0,
        width=48.0,
        height=15.5,
        weight=110.0,
        is_stackable=False,
        stack_qty=1,
    ),
    "half_pallet": ItemPreset(
        name="Half Pallet",
        length=20.0,
        width=48.0,
        height=52.7,
        weight=350.0,
        is_stackable=False,
        stack_qty=1,
    ),
    "quarter_pallet_20x20": ItemPreset(
        name="Quarter Pallet 20x20",
        length=20.0,
        width=20.0,
        height=0.0,
        weight=0.0,
        is_stackable=False,
        stack_qty=1,
    ),
    "quarter_pallet_24x24": ItemPreset(
        name="Quarter Pallet 24x24",
        length=24.0,
        width=24.0,
        height=0.0,
        weight=0.0,
        is_stackable=False,
        stack_qty=1,
    ),
    "custom": ItemPreset(
        name="Custom",
        length=0.0,
        width=0.0,
        height=0.0,
        weight=0.0,
        is_stackable=False,
        stack_qty=1,
    ),
}


ITEM_TYPE_BY_ITEM_NUMBER = {
    "681943924": "full_pallet",
    "683054545": "half_pallet",
}

ITEM_COLOR_BY_ITEM_NUMBER = {
    "681943924": "#4ECDC4",  # PURE teal
    "683054545": "#FF6B9A",  # CDW pink
}

UNKNOWN_ITEM_COLOR = "#98A2B3"

# Color palette for item numbers (cycling)
LOAD_GROUP_COLORS = [
    "#FF6B6B",  # Red
    "#4ECDC4",  # Teal
    "#45B7D1",  # Blue
    "#FFA07A",  # Light salmon
    "#98D8C8",  # Mint
    "#F7DC6F",  # Yellow
    "#BB8FCE",  # Purple
    "#85C1E9",  # Light blue
    "#F8B88B",  # Peach
    "#82E0AA",  # Green
]


def get_preset(preset_key: str) -> TruckPreset:
    """Get a truck preset by key, default to Dry Van."""
    return TRUCK_PRESETS.get(preset_key, TRUCK_PRESETS["dry_van"])


def get_load_group_color(load_group: str, group_index: int) -> str:
    """Get a color for a legacy load group."""
    color_index = group_index % len(LOAD_GROUP_COLORS)
    return LOAD_GROUP_COLORS[color_index]


def get_color_for_group(load_group: str, group_map: dict[str, int]) -> str:
    """Get color for a specific load group using a group-to-index mapping."""
    if load_group not in group_map:
        group_map[load_group] = len(group_map)
    return get_load_group_color(load_group, group_map[load_group])


def get_color_for_item(item_number: str, item_map: dict[str, int]) -> str:
    """Get a stable color for an item number during one planning run."""
    item_key = item_number or "UNKNOWN"
    if item_key in ITEM_COLOR_BY_ITEM_NUMBER:
        return ITEM_COLOR_BY_ITEM_NUMBER[item_key]
    if item_key not in item_map:
        item_map[item_key] = len(item_map)
    return LOAD_GROUP_COLORS[item_map[item_key] % len(LOAD_GROUP_COLORS)]
