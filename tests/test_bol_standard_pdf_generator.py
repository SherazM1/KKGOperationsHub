from __future__ import annotations

from pathlib import Path

from app.models.bol_standard_record import (
    BolAddressBlock,
    BolStandardItemLine,
    BolStandardRecord,
)
from app.services.bol_file_bundle_service import create_standard_bundles
from app.services.bol_standard_docx_generator import GeneratedDocxFile
from app.services.bol_standard_pdf_converter import StandardPdfConversionResult
import app.services.bol_standard_pdf_generator as pdf_generator
from app.services.bol_standard_pdf_generator import generate_standard_pdf_set
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP, BOL_FACILITY_OPTIONS


def _ready_record() -> BolStandardRecord:
    return BolStandardRecord(
        bol_number="10001859231-0553",
        ship_date="2026-05-13",
        carrier="Test Carrier",
        kk_load_number="1",
        kk_po_number="KKPO-001",
        po_number="10001859231-0553",
        dc_number="0553",
        consignee_company="Test DC",
        consignee_street="123 Test Street",
        consignee_city_state_zip="Dallas, TX 75001",
        ship_from=BolAddressBlock(
            company="Kendal King C/O Shorr",
            street="981 W Oakdale Rd",
            city_state_zip="Grand Prairie, TX 75056",
        ),
        bill_to=BolAddressBlock(
            company="Trident Transport, LLC",
            street="505 Riverfront Pkwy",
            city_state_zip="Chattanooga, TN 37402",
        ),
        seal_number_blank="",
        comments="Handle cleanly",
        item_lines=[
            BolStandardItemLine(
                source_row_number=2,
                pallet_qty="2",
                type="PLT",
                po_number="10001859231-0553",
                item_description="Test pallet",
                item_number="ITEM1",
                upc="000111222333",
                skids="2",
                weight_each="100",
                total_weight="306",
            )
        ],
        total_skids=2,
        is_ready=True,
        status="Ready",
        carrier_pro_number="1073839",
        pickup_number="PU-123",
    )


def _generated_docx_file(tmp_path: Path, mode_prefix: str = "standard_bol") -> GeneratedDocxFile:
    source_docx = tmp_path / f"{mode_prefix}_10001859231-0553.docx"
    source_docx.write_bytes(b"placeholder docx metadata source")
    return GeneratedDocxFile(
        bol_number="10001859231-0553",
        file_name=source_docx.name,
        file_path=str(source_docx),
    )


def test_generate_standard_pdf_set_creates_compatible_result_and_bundle(tmp_path: Path) -> None:
    docx_file = _generated_docx_file(tmp_path)

    result = generate_standard_pdf_set(
        [_ready_record()],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[docx_file],
        mode="Standard",
        bol_type="PLT",
        qty_type="Case",
        output_dir=tmp_path / "pdf",
    )

    assert isinstance(result, StandardPdfConversionResult)
    assert result.conversion_available is True
    assert result.converter_name == "reportlab-direct"
    assert result.converted_count == 1
    assert result.failed_count == 0
    assert Path(result.converted_files[0].file_path).exists()
    assert Path(result.converted_files[0].file_path).suffix == ".pdf"

    bundle_result = create_standard_bundles(
        generated_docx_files=[docx_file],
        converted_pdf_files=result.converted_files,
        output_dir=tmp_path / "bundles",
        bundle_name_prefix="standard_bol",
    )

    assert bundle_result.pdf_bundle is not None
    assert bundle_result.pdf_bundle.file_count == 1
    assert bundle_result.all_files_bundle is not None
    assert bundle_result.all_files_bundle.file_count == 2


def test_generate_no_recourse_pdf_set_creates_one_pdf_per_record(tmp_path: Path) -> None:
    docx_file = _generated_docx_file(tmp_path, mode_prefix="no_recourse_bol")

    result = generate_standard_pdf_set(
        [_ready_record()],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[docx_file],
        mode="No Recourse",
        bol_type="CASE",
        qty_type="PLT",
        output_dir=tmp_path / "pdf",
    )

    assert result.converted_count == 1
    assert result.failed_count == 0
    assert result.converted_files[0].file_name == "no_recourse_bol_10001859231-0553.pdf"
    assert Path(result.converted_files[0].file_path).stat().st_size > 0


def test_generate_standard_pdf_set_reports_missing_matching_record(tmp_path: Path) -> None:
    docx_file = GeneratedDocxFile(
        bol_number="MISSING",
        file_name="standard_bol_missing.docx",
        file_path=str(tmp_path / "standard_bol_missing.docx"),
    )

    result = generate_standard_pdf_set(
        [_ready_record()],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[docx_file],
        mode="Standard",
        output_dir=tmp_path / "pdf",
    )

    assert result.converted_count == 0
    assert result.failed_count == 1
    assert "Matching BOL record" in result.failed_conversions[0].error


def test_generate_standard_pdf_uses_carrier_pro_and_unsplit_case_type(tmp_path: Path, monkeypatch) -> None:
    right_fields: list[tuple[str, str]] = []
    table_values: list[str] = []
    original_right_field = pdf_generator._draw_right_field
    original_table_cell = pdf_generator._draw_table_cell_text

    def capture_right_field(canv, row, label, value, **kwargs):
        right_fields.append((label, value))
        return original_right_field(canv, row, label, value, **kwargs)

    def capture_table_cell(canv, col_start, col_end, row_start, row_end, text, **kwargs):
        if col_start == 1 and col_end == 3:
            table_values.append(text)
        return original_table_cell(
            canv,
            col_start,
            col_end,
            row_start,
            row_end,
            text,
            **kwargs,
        )

    monkeypatch.setattr(pdf_generator, "_draw_right_field", capture_right_field)
    monkeypatch.setattr(pdf_generator, "_draw_table_cell_text", capture_table_cell)

    result = generate_standard_pdf_set(
        [_ready_record()],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[_generated_docx_file(tmp_path)],
        mode="Standard",
        bol_type="CASE",
        qty_type="PLT",
        output_dir=tmp_path / "pdf",
    )

    assert result.failed_count == 0
    assert ("Carrier Pro #", "1073839") in right_fields
    assert ("KKG Load #", "1") in right_fields
    assert "CASE" in table_values
    assert "C A S E" not in table_values
    assert "CAS\nE" not in table_values
