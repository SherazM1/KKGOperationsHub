from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pandas as pd
import pytest

from app.models.bol_multistop_row import BolMultistopRow
from app.services.bol_file_bundle_service import create_multistop_docx_bundle
from app.services.bol_multistop_docx_generator import (
    MultistopGeneratedDocxFile,
    generate_multistop_docx_set,
)
from app.services.bol_multistop_mapper import map_multistop_rows_to_records
from app.services.bol_multistop_parser import (
    _parse_multistop_dataframe_rows,
    _resolve_columns,
    _resolve_optional_columns,
)


def _row(
    *,
    kk_load: str,
    stop: int,
    bol_number: str,
    load_number: str = "",
    source_row_number: int | None = None,
) -> BolMultistopRow:
    return BolMultistopRow(
        source_row_number=source_row_number or stop + 1,
        kk_load=kk_load,
        stop=str(stop),
        stop_number=stop,
        trackers=f"TRK-{kk_load}-{stop}",
        carrier="Carrier",
        load_number=load_number,
        kk_po_number=f"KKPO-{kk_load}",
        bol_number=bol_number,
        ship_date="2026-05-13",
        dc_name=f"DC {stop}",
        dc_address=f"{stop} Test St",
        dc_city_state_zip=f"City {stop}, TX 7500{stop}",
        dc_city=f"City {stop}",
        dc_state="TX",
        dc_zip=f"7500{stop}",
        dc_number=f"DC{stop}",
        target_po_number=f"TGT-{stop}",
        item_number=f"ITEM-{stop}",
        upc=f"UPC-{stop}",
        pallet_description=f"Pallet {stop}",
        cases="10",
        total_pallets="2",
        kit_value_each="",
        shipment_value="",
        chargeback_3_percent="",
        weight_each="100",
        weight="100",
    )


def _fake_multistop_saves(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, list[int]]]:
    calls: list[tuple[str, list[int]]] = []

    def fake_combined(**kwargs):
        record = kwargs["record"]
        destination = Path(kwargs["output_root"]) / f"{kwargs['base_name']}.docx"
        destination.write_bytes(b"combined")
        calls.append(("combined", [stop.stop_number for stop in record.stops]))
        return MultistopGeneratedDocxFile(
            bol_number=kwargs["bol_label"],
            file_name=destination.name,
            file_path=str(destination),
            document_type="combined",
            load_number=record.load_number,
            kk_load_number=record.kk_load_number,
            stop_number=None,
        )

    def fake_stop(**kwargs):
        record = kwargs["record"]
        stop = record.stops[kwargs["stop_index"]]
        destination = Path(kwargs["output_root"]) / f"{kwargs['base_name']}.docx"
        destination.write_bytes(b"stop")
        calls.append(("stop", [stop.stop_number]))
        return MultistopGeneratedDocxFile(
            bol_number=stop.bol_number,
            file_name=destination.name,
            file_path=str(destination),
            document_type="stop",
            load_number=record.load_number,
            kk_load_number=record.kk_load_number,
            stop_number=stop.stop_number,
        )

    monkeypatch.setattr(
        "app.services.bol_multistop_docx_generator._save_multistop_docx",
        fake_combined,
    )
    monkeypatch.setattr(
        "app.services.bol_multistop_docx_generator._save_individual_stop_docx",
        fake_stop,
    )
    return calls


def test_multistop_groups_three_stops_by_kk_load_and_generates_one_combined(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    records = map_multistop_rows_to_records(
        [
            _row(kk_load="1", stop=1, bol_number="A", load_number="L1"),
            _row(kk_load="1", stop=2, bol_number="B", load_number="L2"),
            _row(kk_load="1", stop=3, bol_number="C", load_number=""),
        ]
    )
    calls = _fake_multistop_saves(monkeypatch)

    result = generate_multistop_docx_set(
        records,
        selected_facility={"facility_name": "Test", "address": "A", "location": "B"},
        output_dir=tmp_path / "docx",
    )

    assert len(records) == 1
    assert records[0].kk_load_number == "1"
    assert records[0].group_key == "kk_load::1"
    assert [stop.stop_number for stop in records[0].stops] == [1, 2, 3]
    assert [stop.bol_number for stop in records[0].stops] == ["A", "B", "C"]
    assert len([file for file in result.generated_files if file.document_type == "stop"]) == 3
    assert len([file for file in result.generated_files if file.document_type == "combined"]) == 1
    assert calls[0] == ("combined", [1, 2, 3])

    bundle = create_multistop_docx_bundle(result.generated_files, output_dir=tmp_path / "bundle")
    assert bundle.docx_bundle is not None
    with ZipFile(bundle.docx_bundle.file_path) as zip_file:
        names = sorted(zip_file.namelist())

    assert names == [
        "KK_Load_1/combined_multistop_bol_KK_Load_1.docx",
        "KK_Load_1/stop_1_bol_A.docx",
        "KK_Load_1/stop_2_bol_B.docx",
        "KK_Load_1/stop_3_bol_C.docx",
    ]


def test_multistop_groups_multiple_loads_without_cross_contamination() -> None:
    records = map_multistop_rows_to_records(
        [
            _row(kk_load="1", stop=1, bol_number="A"),
            _row(kk_load="1", stop=2, bol_number="B"),
            _row(kk_load="1", stop=3, bol_number="C"),
            _row(kk_load="2", stop=1, bol_number="D"),
            _row(kk_load="2", stop=2, bol_number="E"),
        ]
    )

    assert len(records) == 2
    by_load = {record.kk_load_number: record for record in records}
    assert [stop.stop_number for stop in by_load["1"].stops] == [1, 2, 3]
    assert [stop.bol_number for stop in by_load["1"].stops] == ["A", "B", "C"]
    assert [stop.stop_number for stop in by_load["2"].stops] == [1, 2]
    assert [stop.bol_number for stop in by_load["2"].stops] == ["D", "E"]


def test_multistop_unique_bol_numbers_do_not_split_kk_load_group() -> None:
    records = map_multistop_rows_to_records(
        [
            _row(kk_load="1", stop=1, bol_number="A"),
            _row(kk_load="1", stop=2, bol_number="B"),
            _row(kk_load="1", stop=3, bol_number="C"),
        ]
    )

    assert len(records) == 1
    assert [stop.bol_number for stop in records[0].stops] == ["A", "B", "C"]


def test_multistop_load_number_column_does_not_control_grouping() -> None:
    records = map_multistop_rows_to_records(
        [
            _row(kk_load="1", stop=1, bol_number="A", load_number=""),
            _row(kk_load="1", stop=2, bol_number="B", load_number="DIFFERENT"),
            _row(kk_load="1", stop=3, bol_number="C", load_number="OTHER"),
        ]
    )

    assert len(records) == 1
    assert records[0].kk_load_number == "1"
    assert records[0].load_number == "DIFFERENT"
    assert [stop.stop_number for stop in records[0].stops] == [1, 2, 3]


def test_multistop_stops_are_ordered_numerically_within_kk_load() -> None:
    records = map_multistop_rows_to_records(
        [
            _row(kk_load="1", stop=3, bol_number="C", source_row_number=2),
            _row(kk_load="1", stop=1, bol_number="A", source_row_number=3),
            _row(kk_load="1", stop=2, bol_number="B", source_row_number=4),
        ]
    )

    assert [stop.stop_number for stop in records[0].stops] == [1, 2, 3]
    assert [stop.bol_number for stop in records[0].stops] == ["A", "B", "C"]


def test_halloween_workbook_multistop_grouping_by_kk_load_if_available() -> None:
    workbook_path = Path(r"C:\Users\shera\Downloads\Load_Sheet_halloween_2026_burts.xlsx")
    if not workbook_path.exists():
        pytest.skip("Halloween Multistop load sheet fixture is not available locally.")

    workbook = pd.ExcelFile(workbook_path)
    sheet_name = workbook.sheet_names[0]
    df = workbook.parse(sheet_name=sheet_name, dtype=object)
    column_map = _resolve_columns(df.columns.tolist(), sheet_name)
    optional_column_map = _resolve_optional_columns(df.columns.tolist())
    rows = _parse_multistop_dataframe_rows(df, column_map, optional_column_map)
    records = map_multistop_rows_to_records(rows)

    assert len(records) == 12
    by_load = {record.kk_load_number: record for record in records}
    for load_number in ("1", "5", "6", "7", "11", "12"):
        assert len(by_load[load_number].stops) == 3
