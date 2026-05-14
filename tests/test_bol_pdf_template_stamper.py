from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader

from app.models.bol_multistop_record import BolMultistopRecord, BolMultistopStop
from app.models.bol_standard_record import (
    BolAddressBlock,
    BolStandardItemLine,
    BolStandardRecord,
)
from app.services.bol_file_bundle_service import create_standard_bundles
from app.services.bol_multistop_docx_generator import MultistopGeneratedDocxFile
from app.services.bol_pdf_template_stamper import stamp_bol_pdf_set
from app.services.bol_standard_docx_generator import GeneratedDocxFile, StandardDocxGenerationResult
from app.services.bol_standard_pdf_converter import StandardPdfConversionResult
from app.ui import bol_generator
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP, BOL_FACILITY_OPTIONS


def _address(company: str = "Trident Transport, LLC") -> BolAddressBlock:
    return BolAddressBlock(
        company=company,
        street="505 Riverfront Pkwy",
        city_state_zip="Chattanooga, TN 37402",
    )


def _standard_record(
    *,
    optional_fields: bool = True,
    bol_number: str = "10001859231-0553",
    po_number: str = "10001859231-0553",
) -> BolStandardRecord:
    return BolStandardRecord(
        bol_number=bol_number,
        ship_date="2026-05-13",
        carrier="Test Carrier",
        kk_load_number="1",
        kk_po_number="KKPO-001",
        po_number=po_number,
        dc_number="0553",
        consignee_company="Test DC",
        consignee_street="123 Test Street",
        consignee_city_state_zip="Dallas, TX 75001",
        ship_from=_address("Kendal King C/O Shorr"),
        bill_to=_address(),
        seal_number_blank="SEAL-1" if optional_fields else "",
        comments="Handle cleanly" if optional_fields else "",
        item_lines=[
            BolStandardItemLine(
                source_row_number=2,
                pallet_qty="2",
                type="PLT",
                po_number=po_number,
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
        pickup_number="PU-123" if optional_fields else "",
    )


def _multistop_record() -> BolMultistopRecord:
    stops = [
        BolMultistopStop(
            source_row_number=2,
            stop_number=1,
            delivery_dc="DC 0551",
            delivery_address="1 Stop Way",
            delivery_city_state_zip="Dallas, TX 75001",
            dc_number="0551",
            cases="10",
            target_po_number="PO-0551",
            pallet_description="Stop one pallet",
            item_number="ITEM1",
            upc="000111222333",
            total_pallets="2",
            weight="100",
        ),
        BolMultistopStop(
            source_row_number=3,
            stop_number=2,
            delivery_dc="DC 0553",
            delivery_address="2 Stop Way",
            delivery_city_state_zip="Austin, TX 73301",
            dc_number="0553",
            cases="20",
            target_po_number="PO-0553",
            pallet_description="Stop two pallet",
            item_number="ITEM2",
            upc="000111222334",
            total_pallets="3",
            weight="200",
        ),
    ]
    return BolMultistopRecord(
        bol_number="MBOL-001",
        ship_date="2026-05-13",
        carrier="Test Carrier",
        load_number="LOAD-001",
        kk_po_number="KKPO-001",
        kk_load_number="1",
        group_key="MBOL-001::LOAD-001",
        stop_count=2,
        stops=stops,
        delivery_1_dc="DC 0551",
        delivery_1_address="1 Stop Way\nDallas, TX 75001",
        delivery_2_dc="DC 0553",
        delivery_2_address="2 Stop Way\nAustin, TX 73301",
        delivery_3_dc="",
        delivery_3_address="",
        dc_1="0551",
        case_1="10",
        po_1="PO-0551",
        pallet_description_1="Stop one pallet",
        plt_1="2",
        weight_1="100",
        dc_2="0553",
        case_2="20",
        po_2="PO-0553",
        pallet_description_2="Stop two pallet",
        plt_2="3",
        weight_2="200",
        dc_3="",
        case_3="",
        po_3="",
        pallet_description_3="",
        plt_3="",
        weight_3="",
        total_case=30,
        total_pallet=5,
        total_ship_weight=300,
        po_number="PO-0551",
        dc_number="0551",
        consignee_company="DC 0551",
        consignee_street="1 Stop Way",
        consignee_city_state_zip="Dallas, TX 75001",
        ship_from=_address("Kendal King C/O Shorr"),
        bill_to=_address(),
        seal_number_blank="",
        comments="",
        item_lines=[],
        total_skids=30,
        is_ready=True,
        status="Ready",
    )


def _docx_file(tmp_path: Path, name: str, bol_number: str) -> GeneratedDocxFile:
    source = tmp_path / name
    source.write_bytes(b"placeholder")
    return GeneratedDocxFile(
        bol_number=bol_number,
        file_name=source.name,
        file_path=str(source),
    )


def _pdf_text(path: str) -> str:
    reader = PdfReader(path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def test_standard_template_stamper_creates_pdf_and_bundle(tmp_path: Path) -> None:
    docx_file = _docx_file(tmp_path, "standard_bol_10001859231-0553.docx", "10001859231-0553")

    result = stamp_bol_pdf_set(
        [_standard_record()],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[docx_file],
        mode="Standard",
        bol_type="CASE",
        qty_type="PLT",
        output_dir=tmp_path / "pdf",
    )

    assert isinstance(result, StandardPdfConversionResult)
    assert result.converter_name == "pdf-template-stamper"
    assert result.converted_count == 1
    assert result.failed_count == 0
    assert Path(result.converted_files[0].file_path).exists()

    text = _pdf_text(result.converted_files[0].file_path)
    assert "1073839" in text
    assert "CASE" in text
    assert "Pallet Qty" in text
    assert "C A S E" not in text
    assert "CAS\nE" not in text

    bundle = create_standard_bundles(
        generated_docx_files=[docx_file],
        converted_pdf_files=result.converted_files,
        output_dir=tmp_path / "bundles",
    )
    assert bundle.pdf_bundle is not None
    assert bundle.pdf_bundle.file_count == 1


def test_no_recourse_template_stamper_creates_one_page_pdf_with_missing_optional_fields(tmp_path: Path) -> None:
    docx_file = _docx_file(tmp_path, "no_recourse_bol_10001859231-0553.docx", "10001859231-0553")

    result = stamp_bol_pdf_set(
        [_standard_record(optional_fields=False)],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[docx_file],
        mode="No Recourse",
        bol_type="CASE",
        qty_type="Case",
        output_dir=tmp_path / "pdf",
    )

    assert result.converted_count == 1
    assert result.failed_count == 0
    reader = PdfReader(result.converted_files[0].file_path)
    assert len(reader.pages) == 1
    text = _pdf_text(result.converted_files[0].file_path)
    assert "1073839" in text
    assert "CASE" in text
    assert "Case Qty" in text


def test_no_recourse_template_stamper_keeps_long_bol_and_po_values_complete(tmp_path: Path) -> None:
    long_value = "10001859231-0553-EXTRA"
    docx_file = _docx_file(tmp_path, f"no_recourse_bol_{long_value}.docx", long_value)

    result = stamp_bol_pdf_set(
        [_standard_record(bol_number=long_value, po_number=long_value)],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[docx_file],
        mode="No Recourse",
        bol_type="CASE",
        qty_type="Case",
        output_dir=tmp_path / "pdf",
    )

    assert result.converted_count == 1
    text = _pdf_text(result.converted_files[0].file_path)
    assert long_value in text
    assert "CASE" in text
    assert "C A S E" not in text


def test_multistop_template_stamper_creates_pdf(tmp_path: Path) -> None:
    source = tmp_path / "combined_multistop_bol_MBOL-001_LOAD-001.docx"
    source.write_bytes(b"placeholder")
    docx_file = MultistopGeneratedDocxFile(
        bol_number="MBOL-001",
        file_name=source.name,
        file_path=str(source),
        document_type="combined",
        load_number="LOAD-001",
        stop_number=None,
    )

    result = stamp_bol_pdf_set(
        [_multistop_record()],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        generated_docx_files=[docx_file],
        mode="Multistop",
        output_dir=tmp_path / "pdf",
    )

    assert result.converted_count == 1
    assert result.failed_count == 0
    assert Path(result.converted_files[0].file_path).exists()
    text = _pdf_text(result.converted_files[0].file_path)
    assert "MBOL-001" in text
    assert "LOAD-001" in text


def test_ui_pdf_generation_routes_supported_modes_to_template_stamper(monkeypatch, tmp_path: Path) -> None:
    calls: list[str] = []
    docx_file = _docx_file(tmp_path, "standard_bol_10001859231-0553.docx", "10001859231-0553")
    docx_result = StandardDocxGenerationResult(
        output_dir=str(tmp_path),
        generated_files=[docx_file],
        skipped_records=[],
        failed_records=[],
        notices=[],
    )

    def fake_stamp(records, selected_facility, generated_docx_files, *, mode, **kwargs):
        calls.append(mode)
        return StandardPdfConversionResult(
            output_dir=str(tmp_path / mode),
            converted_files=[],
            failed_conversions=[],
            converter_name="pdf-template-stamper",
            conversion_available=True,
            unavailable_reason=None,
            converter_path=None,
        )

    monkeypatch.setattr(bol_generator, "stamp_bol_pdf_set", fake_stamp)
    bol_generator.st.session_state["bol_selected_facility"] = BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]]
    bol_generator.st.session_state["bol_type_selector"] = "CASE"
    bol_generator.st.session_state["bol_qty_type_selector"] = "PLT"
    bol_generator.st.session_state["bol_batch_comment_textarea"] = ""

    for mode in ("Standard", "No Recourse", "Multistop"):
        bol_generator._generate_pdf_result(
            mode=mode,
            docx_result=docx_result,
            grouped_records=[],
            progress_callback=None,
        )

    assert calls == ["Standard", "No Recourse", "Multistop"]
