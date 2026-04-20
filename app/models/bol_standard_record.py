"""Grouped Standard BOL record models for downstream generation."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class BolAddressBlock:
    """Address block used by ship-from and bill-to sections."""

    company: str
    street: str
    city_state_zip: str
    attn: str = ""


@dataclass(slots=True)
class BolStandardItemLine:
    """One mapped item line under a grouped Standard BOL record."""

    source_row_number: int
    pallet_qty: str
    type: str
    po_number: str
    item_description: str
    item_number: str
    upc: str
    skids: str
    weight_each: str


@dataclass(slots=True)
class BolStandardRecord:
    """One grouped Standard BOL record keyed by BOL #."""

    bol_number: str
    ship_date: str
    carrier: str
    kk_load_number: str
    kk_po_number: str
    po_number: str
    dc_number: str
    consignee_company: str
    consignee_street: str
    consignee_city_state_zip: str
    ship_from: BolAddressBlock
    bill_to: BolAddressBlock
    seal_number_blank: str
    comments: str
    item_lines: list[BolStandardItemLine]
    total_skids: float
    is_ready: bool
    status: str
    selected_for_generation: bool = True
    missing_required_fields: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    generation_skip_reason: str | None = None
    conversion_skip_reason: str | None = None
    issues: list[str] = field(default_factory=list)
