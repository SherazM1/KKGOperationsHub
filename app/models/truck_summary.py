"""Truck assignment and summary model."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class BoxLayout:
    """Represents a single load box/pallet in a truck."""

    box_id: int
    load_group: str
    pallet_count: float
    description: str
    color: str
    x: float  # position in truck (normalized 0-1)
    y: float  # position in truck (normalized 0-1)
    width: float  # size (normalized 0-1)
    height: float  # size (normalized 0-1)


@dataclass
class TruckSummary:
    """Represents a truck assignment and its contents."""

    truck_id: int
    truck_preset: str
    total_capacity_pallets: float
    total_used_pallets: float
    utilization_percent: float
    remaining_capacity: float
    load_groups: list[str] = field(default_factory=list)
    items_count: int = 0
    boxes: list[BoxLayout] = field(default_factory=list)
    notes: str = ""

    @property
    def utilization_display(self) -> str:
        """Format utilization for display."""
        return f"{self.utilization_percent:.1f}%"
