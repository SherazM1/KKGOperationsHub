from __future__ import annotations

from pathlib import Path

from docx import Document

from app.models.bol_standard_record import (
    BolAddressBlock,
    BolStandardItemLine,
    BolStandardRecord,
)
from app.services.bol_standard_docx_generator import (
    generate_standard_docx_set,
    resolve_output_filename_prefix_for_mode,
    resolve_template_path_for_mode,
)
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP, BOL_FACILITY_OPTIONS


def _ready_record() -> BolStandardRecord:
    return BolStandardRecord(
        bol_number="10001859231-0553",
        ship_date="2026-05-13",
        carrier="Test Carrier",
        kk_load_number="1073839",
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
        comments="",
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
            )
        ],
        total_skids=2,
        is_ready=True,
        status="Ready",
    )


def _generated_docx(mode: str, tmp_path: Path, *, bol_type: str, qty_type: str) -> Document:
    result = generate_standard_docx_set(
        [_ready_record()],
        selected_facility=BOL_FACILITY_LOOKUP[BOL_FACILITY_OPTIONS[0]],
        bol_type=bol_type,
        qty_type=qty_type,
        template_path=resolve_template_path_for_mode(mode),
        output_dir=tmp_path,
        file_name_prefix=resolve_output_filename_prefix_for_mode(mode),
    )

    assert result.generated_count == 1
    return Document(result.generated_files[0].file_path)


def _item_header_text(doc: Document) -> str:
    table = doc.tables[0]
    for row in table.rows:
        row_text = " ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
        row_text_upper = row_text.upper()
        if "TYPE" in row_text_upper and "PO #" in row_text_upper and "ITEM DESCRIPTION" in row_text_upper:
            return row.cells[0].text.strip()
    raise AssertionError("Item table header row was not found.")


def _first_item_type_value(doc: Document) -> str:
    table = doc.tables[0]
    for row in table.rows:
        row_text = " ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
        if "Test pallet" in row_text:
            return row.cells[1].text.strip()
    raise AssertionError("First item row was not found.")


def test_standard_docx_qty_type_plt_renders_pallet_qty_header(tmp_path: Path) -> None:
    doc = _generated_docx("Standard", tmp_path, bol_type="CASE", qty_type="PLT")

    assert _item_header_text(doc) == "Pallet Qty"
    assert _first_item_type_value(doc) == "CASE"


def test_standard_docx_qty_type_case_renders_case_qty_header(tmp_path: Path) -> None:
    doc = _generated_docx("Standard", tmp_path, bol_type="PLT", qty_type="Case")

    assert _item_header_text(doc) == "Case Qty"
    assert _first_item_type_value(doc) == "PLT"


def test_no_recourse_docx_qty_type_case_renders_case_qty_header(tmp_path: Path) -> None:
    doc = _generated_docx("No Recourse", tmp_path, bol_type="PLT", qty_type="Case")

    assert _item_header_text(doc) == "Case Qty"
    assert _first_item_type_value(doc) == "PLT"
