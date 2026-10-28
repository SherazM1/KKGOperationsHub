"""Truck assignment and summary model."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class BoxLayout:
    """Represents one visual item box in a truck."""

    box_id: int
    load_group: str
    pallet_count: float
    description: str
    color: str
    x: float  # position in truck (normalized 0-1)
    y: float  # position in truck (normalized 0-1)
    width: float  # size (normalized 0-1)
    height: float  # size (normalized 0-1)
    kkg_load_number: str = ""
    retailer_po_number: str = ""
    item_number: str = ""
    row_qty: int = 1
    represented_qty: int = 1
    unit_index: int = 1
    item_length: float = 0.0
    item_width: float = 0.0
    item_height: float = 0.0
    grouped_height: float = 0.0
    item_weight: float = 0.0
    stack_level: int = 1
    stack_qty: int = 1
    item_type: str = ""
    store_number: str = ""
    dc_number: str = ""
    stack_members: list["BoxLayout"] = field(default_factory=list)

    @property
    def allocations(self) -> list["BoxLayout"]:
        """Source quantities represented by this floor position."""
        return self.stack_members or [self]


@dataclass
class TruckSummary:
    """Represents a single KKG load assignment and its truck fit status."""

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
    kkg_load_number: str = ""
    truck_length: float = 0.0
    truck_width: float = 0.0
    truck_height: float = 0.0
    truck_max_weight: float = 0.0
    total_weight: float = 0.0
    weight_utilization_percent: float = 0.0
    spatial_fit: bool = True
    weight_fit: bool = True
    validation_status: str = "pending"
    validation_notes: list[str] = field(default_factory=list)
    pallet_positions_used: float | None = None
    pallet_positions_capacity: float | None = None
    roof_clearance_in: float = 3.0

    @property
    def utilization_display(self) -> str:
        """Format utilization for display."""
        return f"{self.utilization_percent:.1f}%"

    @property
    def load_capacity_display(self) -> str:
        if self.pallet_positions_capacity:
            return f"{(self.pallet_positions_used or 0) / self.pallet_positions_capacity * 100:.0f}%"
        return self.utilization_display
