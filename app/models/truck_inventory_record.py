"""Normalized truck inventory record model."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class TruckInventoryRecord:
    """Represents a normalized inventory item from PURE or CDW."""

    source_type: str  # "pure" or "cdw"
    source_file: str
    po_number: str
    po_date: str | None = None
    event_code: str | None = None
    delivery_date: str | None = None
    
    dc_number: str | None = None
    dc_name: str | None = None
    receiver_name: str | None = None
    
    ship_to_address_1: str | None = None
    ship_to_address_2: str | None = None
    ship_to_city: str | None = None
    ship_to_state: str | None = None
    ship_to_zip: str | None = None
    
    item_number: str | None = None
    upc: str | None = None
    description: str | None = None
    qty: int | None = None
    uom: str | None = None
    unit_weight: float | None = None
    total_weight: float | None = None
    
    pallet_description: str | None = None
    pallet_type: str | None = None
    estimated_pallets: float | None = None
    
    load_group: str | None = None
    validation_status: str = "pending"  # "pending", "valid", "warning", "error"
    validation_notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        """Convert record to dictionary for DataFrame conversion."""
        return {
            "source_type": self.source_type,
            "source_file": self.source_file,
            "po_number": self.po_number,
            "po_date": self.po_date,
            "event_code": self.event_code,
            "delivery_date": self.delivery_date,
            "dc_number": self.dc_number,
            "dc_name": self.dc_name,
            "receiver_name": self.receiver_name,
            "ship_to_address_1": self.ship_to_address_1,
            "ship_to_address_2": self.ship_to_address_2,
            "ship_to_city": self.ship_to_city,
            "ship_to_state": self.ship_to_state,
            "ship_to_zip": self.ship_to_zip,
            "item_number": self.item_number,
            "upc": self.upc,
            "description": self.description,
            "qty": self.qty,
            "uom": self.uom,
            "unit_weight": self.unit_weight,
            "total_weight": self.total_weight,
            "pallet_description": self.pallet_description,
            "pallet_type": self.pallet_type,
            "estimated_pallets": self.estimated_pallets,
            "load_group": self.load_group,
            "validation_status": self.validation_status,
            "validation_notes": ";".join(self.validation_notes),
        }
