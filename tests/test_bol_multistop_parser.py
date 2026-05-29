from __future__ import annotations

from io import BytesIO

import pandas as pd

from app.services.bol_multistop_parser import parse_multistop_bol_excel


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


def test_parse_multistop_bol_excel_accepts_csv_upload() -> None:
    csv_file = _csv_with_rows([_multistop_load_row()])

    rows = parse_multistop_bol_excel(csv_file)

    assert len(rows) == 1
    assert rows[0].bol_number == "BOL-001"
    assert rows[0].stop_number == 1
    assert rows[0].item_number == "ITEM-001"
    assert rows[0].upc == "000111222333"
