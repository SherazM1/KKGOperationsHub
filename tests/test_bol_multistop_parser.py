from __future__ import annotations

from io import BytesIO

import pandas as pd
import pytest

from app.services.bol_multistop_parser import (
    _resolve_columns,
    parse_multistop_bol_excel,
)


def _multistop_load_row() -> dict[str, object]:
    return {
        "KK Load": "KL-001",
        "Stop": 1,
        "TRACKERS": "TRK-001",
        "Carrier": "Test Carrier",
        "load#": "LOAD-001",
        "KK PO#": "KKPO-001",
        "BOL #": "BOL-001",
        "ship date": "2026-05-13",
        "DC Name": "Test DC",
        "DC ADDRESS": "123 Test St",
        "DC City, State, Zip": "Dallas, TX 75001",
        "DC CITY": "Dallas",
        "DCST": "TX",
        "DCZIP": "75001",
        "DC #": "1234",
        "COUNTRY": "US",
        "DEPT.": "001",
        "TGT PO #": "TGT-001",
        "MABD": "2026-05-20",
        "UPC": "000111222333",
        "PalletDescription": "Test pallet",
        "Cases": 10,
        "Total PLT": 2,
        "Kit Value (EACH)": 5,
        "Shipment Value": 50,
        "3% Chargeback": 1.5,
        "weight each": 50,
        "Weight": 100,
        "ITEM #": "ITEM-001",
    }


def _csv_with_rows(rows: list[dict[str, object]], name: str = "multistop.csv") -> BytesIO:
    output = BytesIO()
    output.write(pd.DataFrame(rows).to_csv(index=False).encode("utf-8"))
    output.seek(0)
    output.name = name
    return output


def _renamed_row(
    row: dict[str, object],
    rename_map: dict[str, str],
    omit: set[str] | None = None,
    extras: dict[str, object] | None = None,
) -> dict[str, object]:
    omitted = omit or set()
    renamed = {
        rename_map.get(key, key): value
        for key, value in row.items()
        if key not in omitted
    }
    if extras:
        renamed.update(extras)
    return renamed


def test_parse_multistop_bol_excel_accepts_csv_upload() -> None:
    csv_file = _csv_with_rows([_multistop_load_row()])

    rows = parse_multistop_bol_excel(csv_file)

    assert len(rows) == 1
    assert rows[0].bol_number == "BOL-001"
    assert rows[0].stop_number == 1
    assert rows[0].item_number == "ITEM-001"
    assert rows[0].upc == "000111222333"


def test_parse_multistop_bol_excel_accepts_capitalized_operational_headers() -> None:
    row = _renamed_row(
        _multistop_load_row(),
        {
            "Stop": "STOP",
            "DCST": "DC ST",
            "DCZIP": "DC ZIP",
            "Cases": "CASE QTY",
            "Total PLT": "TOTAL PALLETS",
        },
        omit={"DC City, State, Zip", "weight each"},
    )
    csv_file = _csv_with_rows([row])

    rows = parse_multistop_bol_excel(csv_file)

    assert rows[0].dc_city_state_zip == "Dallas, TX 75001"
    assert rows[0].dc_state == "TX"
    assert rows[0].dc_zip == "75001"
    assert rows[0].cases == "10"
    assert rows[0].total_pallets == "2"
    assert rows[0].weight_each == "100"


def test_parse_multistop_bol_excel_normalizes_whitespace_and_newlines() -> None:
    row = _renamed_row(
        _multistop_load_row(),
        {
            "DC CITY": " DC CITY ",
            "DCST": "DC\nST",
            "DCZIP": "DC\tZIP",
            "Cases": "Case   Qty",
            "Total PLT": "TOTAL   PALLETS",
        },
        omit={"DC City, State, Zip"},
    )
    csv_file = _csv_with_rows([row])

    rows = parse_multistop_bol_excel(csv_file)

    assert rows[0].dc_city == "Dallas"
    assert rows[0].dc_state == "TX"
    assert rows[0].dc_zip == "75001"
    assert rows[0].cases == "10"
    assert rows[0].total_pallets == "2"


def test_multistop_operational_aliases_resolve_to_canonical_fields() -> None:
    columns = [
        "KK Load",
        "Stop",
        "TRACKERS",
        "Carrier",
        "load#",
        "KK PO#",
        "BOL #",
        "ship date",
        "DC Name",
        "DC ADDRESS",
        "DC CITY",
        "DC ST",
        "DC ZIP",
        "DC #",
        "COUNTRY",
        "DEP.",
        "TGT PO #",
        "MABD",
        "UPC",
        "Pallet Description",
        "Case Qty",
        "Total Pallets",
        "Weight",
    ]

    column_map = _resolve_columns(columns, worksheet_name="CSV")

    assert column_map["dc_city"] == "DC CITY"
    assert column_map["dc_state"] == "DC ST"
    assert column_map["dc_zip"] == "DC ZIP"
    assert column_map["cases"] == "Case Qty"
    assert column_map["total_pallets"] == "Total Pallets"
    assert column_map["weight_each"] == "Weight"


def test_parse_multistop_bol_excel_allows_optional_financial_fields_absent() -> None:
    row = _renamed_row(
        _multistop_load_row(),
        {},
        omit={"Kit Value (EACH)", "Shipment Value", "3% Chargeback"},
    )
    csv_file = _csv_with_rows([row])

    rows = parse_multistop_bol_excel(csv_file)

    assert rows[0].kit_value_each == ""
    assert rows[0].shipment_value == ""
    assert rows[0].chargeback_3_percent == ""


def test_parse_multistop_bol_excel_maps_optional_financial_aliases_when_present() -> None:
    row = _renamed_row(
        _multistop_load_row(),
        {
            "Kit Value (EACH)": "Kit Value Each",
            "Shipment Value": "SHIPMENT VALUE",
            "3% Chargeback": "3 Percent Chargeback",
        },
    )
    csv_file = _csv_with_rows([row])

    rows = parse_multistop_bol_excel(csv_file)

    assert rows[0].kit_value_each == "5"
    assert rows[0].shipment_value == "50"
    assert rows[0].chargeback_3_percent == "1.5"


def test_multistop_weight_alias_does_not_match_total_ship_weight() -> None:
    row = _renamed_row(
        _multistop_load_row(),
        {},
        omit={"weight each"},
        extras={"Total Ship Weight": 9999},
    )
    csv_file = _csv_with_rows([row])

    rows = parse_multistop_bol_excel(csv_file)

    assert rows[0].weight_each == "100"
    assert rows[0].weight == "100"


def test_multistop_parser_still_rejects_truly_missing_required_field() -> None:
    row = _renamed_row(_multistop_load_row(), {}, omit={"Cases"})
    csv_file = _csv_with_rows([row])

    with pytest.raises(ValueError, match="cases"):
        parse_multistop_bol_excel(csv_file)
