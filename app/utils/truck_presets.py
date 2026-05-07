"""Truck presets and configuration for inventory assignments."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class TruckPreset:
    """Represents a standard truck capacity preset."""

    name: str
    pallet_capacity: float  # Standard pallet count capacity
    length_ft: int
    width_ft: int
    height_ft: int
    description: str


# Standard truck presets
TRUCK_PRESETS = {
    "53ft_dry_van": TruckPreset(
        name="53 ft Dry Van",
        pallet_capacity=26.0,
        length_ft=53,
        width_ft=8,
        height_ft=9,
        description="Standard 53ft dry trailer (26 pallets)",
    ),
    "48ft_trailer": TruckPreset(
        name="48 ft Trailer",
        pallet_capacity=24.0,
        length_ft=48,
        width_ft=8,
        height_ft=9,
        description="48ft dry trailer (24 pallets)",
    ),
    "half_truck": TruckPreset(
        name="Half Truck",
        pallet_capacity=13.0,
        length_ft=26,
        width_ft=8,
        height_ft=9,
        description="Half truck (13 pallets)",
    ),
    "quarter_truck": TruckPreset(
        name="Quarter Truck",
        pallet_capacity=6.0,
        length_ft=13,
        width_ft=8,
        height_ft=9,
        description="Quarter truck (6 pallets)",
    ),
}

# Color palette for load groups (cycling)
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
    """Get a truck preset by key, default to 53ft dry van."""
    return TRUCK_PRESETS.get(preset_key, TRUCK_PRESETS["53ft_dry_van"])


def get_load_group_color(load_group: str, group_index: int) -> str:
    """Get a color for a load group."""
    color_index = group_index % len(LOAD_GROUP_COLORS)
    return LOAD_GROUP_COLORS[color_index]


def get_color_for_group(load_group: str, group_map: dict[str, int]) -> str:
    """Get color for a specific load group using a group-to-index mapping."""
    if load_group not in group_map:
        group_map[load_group] = len(group_map)
    return get_load_group_color(load_group, group_map[load_group])
