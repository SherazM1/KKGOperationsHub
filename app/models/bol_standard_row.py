"""Data model for one normalized Standard BOL source row."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class BolStandardRow:
    """Represents one parsed row from the Standard BOL source sheet."""

    source_row_number: int
    bol_number: str
    ship_date: str
    carrier: str
    kk_load: str
    kk_po: str
    wm_po: str
    dc_number: str
    dc_name: str
    dc_street: str
    dc_city_state_zip: str
    item_number: str
    upc: str
    item_description: str
    quantity: str
    weight_each: str

