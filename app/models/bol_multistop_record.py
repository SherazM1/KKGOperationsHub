"""Grouped Multistop BOL record models for downstream generation."""

from __future__ import annotations

from dataclasses import dataclass, field

from app.models.bol_standard_record import BolAddressBlock, BolStandardItemLine


@dataclass(slots=True)
class BolMultistopStop:
    """One mapped stop in a grouped Multistop shipment."""

    source_row_number: int
    stop_number: int
    delivery_dc: str
    delivery_address: str
    delivery_city_state_zip: str
    dc_number: str
    cases: str
    target_po_number: str
    pallet_description: str
    item_number: str
    upc: str
    total_pallets: str
    weight: str


@dataclass(slots=True)
class BolMultistopRecord:
    """One grouped Multistop BOL record keyed by BOL # + load#."""

    # Header and grouping identity.
    bol_number: str
    ship_date: str
    carrier: str
    load_number: str
    kk_po_number: str
    kk_load_number: str
    group_key: str

    # Stop model.
    stop_count: int
    stops: list[BolMultistopStop]

    # Delivery block fields.
    delivery_1_dc: str
    delivery_1_address: str
    delivery_2_dc: str
    delivery_2_address: str
    delivery_3_dc: str
    delivery_3_address: str

    # Line-item fields by stop.
    dc_1: str
    case_1: str
    po_1: str
    pallet_description_1: str
    plt_1: str
    weight_1: str

    dc_2: str
    case_2: str
    po_2: str
    pallet_description_2: str
    plt_2: str
    weight_2: str

    dc_3: str
    case_3: str
    po_3: str
    pallet_description_3: str
    plt_3: str
    weight_3: str

    # Totals.
    total_case: float
    total_pallet: float
    total_ship_weight: float

    # Compatibility fields (shared review/facility/comment flow).
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

    # Readiness/status.
    is_ready: bool
    status: str
    selected_for_generation: bool = True
    missing_required_fields: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    generation_skip_reason: str | None = None
    conversion_skip_reason: str | None = None
    issues: list[str] = field(default_factory=list)
    is_supported: bool = True
