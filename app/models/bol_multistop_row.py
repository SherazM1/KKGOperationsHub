"""Data model for one normalized Multistop BOL source row."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class BolMultistopRow:
    """Represents one parsed row from the Multistop BOL source sheet."""

    source_row_number: int
    kk_load: str
    stop: str
    stop_number: int | None
    trackers: str
    carrier: str
    load_number: str
    kk_po_number: str
    bol_number: str
    ship_date: str
    dc_name: str
    dc_address: str
    dc_city_state_zip: str
    dc_city: str
    dc_state: str
    dc_zip: str
    dc_number: str
    target_po_number: str
    item_number: str
    upc: str
    pallet_description: str
    cases: str
    total_pallets: str
    kit_value_each: str
    shipment_value: str
    chargeback_3_percent: str
    weight_each: str
    weight: str
