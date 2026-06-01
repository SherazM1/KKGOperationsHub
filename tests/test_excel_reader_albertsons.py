"""Tests for Albertsons Excel parsing."""

from __future__ import annotations

from io import BytesIO

import pandas as pd
import pytest

from app.services.excel_reader_albertsons import read_excel_albertsons


def _excel_file(rows: list[dict[str, str]]) -> BytesIO:
    buffer = BytesIO()
    pd.DataFrame(rows).to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer


def _base_row(**overrides: str) -> dict[str, str]:
    row = {
        "Buying Party Name": "ALBERTSONS COMPANIES LLC SUB 123",
        "Buying Party Address 1": "123 Main St",
        "Buying Party City": "Dallas",
        "Buying Party State": "TX",
        "Buying Party Zip": "75001",
        "Purchase Order Number": "PO-1",
        "Item #": "ITEM-1",
        "Description": "Display",
        "Quantity": "12",
    }
    row.update(overrides)
    return row


def test_read_excel_albertsons_normalizes_scientific_upc_text() -> None:
    labels = read_excel_albertsons(
        _excel_file([_base_row(**{"UPC #": "3.98E+09"})]),
        require_upc=True,
    )

    assert labels[0].upc == "3980000000"
    assert labels[0].ship_to_name == "ALBERTSONS COMPANIES LLC"


def test_read_excel_albertsons_preserves_text_upc_leading_zeroes() -> None:
    labels = read_excel_albertsons(
        _excel_file([_base_row(UPC="003980000000")]),
        require_upc=True,
    )

    assert labels[0].upc == "003980000000"


def test_read_excel_albertsons_reads_dc_code_column() -> None:
    labels = read_excel_albertsons(_excel_file([_base_row(**{"DC CODE": "SWCA"})]))

    assert labels[0].dc_label == "DC#"
    assert labels[0].dc_value == "SWCA"


@pytest.mark.parametrize("dc_header", ["DC#", "DC_CODE", "Distribution Center Code"])
def test_read_excel_albertsons_reads_dc_code_aliases(dc_header: str) -> None:
    labels = read_excel_albertsons(_excel_file([_base_row(**{dc_header: "NWCA"})]))

    assert labels[0].dc_value == "NWCA"


def test_read_excel_albertsons_defaults_dc_value_when_column_missing() -> None:
    labels = read_excel_albertsons(_excel_file([_base_row()]))

    assert labels[0].dc_value == "WNCA"


def test_read_excel_albertsons_upc_mode_requires_upc_column() -> None:
    with pytest.raises(
        ValueError,
        match="UPC # from Excel mode requires a UPC # column in the Albertsons Excel file.",
    ):
        read_excel_albertsons(_excel_file([_base_row()]), require_upc=True)
