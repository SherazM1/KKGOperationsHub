"""DOCX generation service for Multistop-mode BOL records."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tempfile import mkdtemp
from xml.sax.saxutils import escape
import zipfile

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips
from docx.table import Table

from app.models.bol_multistop_record import BolMultistopRecord
from app.models.bol_standard_record import BolStandardItemLine, BolStandardRecord
from app.services.bol_standard_docx_generator import (
    DocxGenerationNotice,
    FailedDocxRecord,
    GeneratedDocxFile,
    STANDARD_TEMPLATE_PATH,
    SkippedDocxRecord,
    StandardDocxGenerationResult,
    _apply_template_record_values as _apply_standard_template_record_values,
    _postprocess_comments_in_saved_docx as _postprocess_standard_comments_in_saved_docx,
)
from app.utils.bol_facilities import BolFacilityRecord


MULTISTOP_TEMPLATE_PATH = Path("app/templates/multistop_bol_template.docx")
LEFT_MERGE = "\u00ab"
RIGHT_MERGE = "\u00bb"


@dataclass(slots=True)
class MultistopGeneratedDocxFile(GeneratedDocxFile):
    """Generated Multistop DOCX metadata for combined and stop-level outputs."""

    document_type: str
    load_number: str
    stop_number: int | None = None


def _tok(name: str) -> str:
    return f"{LEFT_MERGE}{name}{RIGHT_MERGE}"


def _sanitize_filename_part(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in ("-", "_") else "_" for char in value)
    cleaned = cleaned.strip("_")
    return cleaned or "unknown"


def _unique_destination_path(directory: Path, base_name: str, extension: str) -> Path:
    candidate = directory / f"{base_name}{extension}"
    if not candidate.exists():
        return candidate

    suffix = 2
    while True:
        candidate = directory / f"{base_name}_{suffix}{extension}"
        if not candidate.exists():
            return candidate
        suffix += 1


def _format_number(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else f"{value:.2f}".rstrip("0").rstrip(".")


def _parse_number(value: str) -> float:
    cleaned = (value or "").replace(",", "").strip()
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _format_ship_date_for_template(raw_ship_date: str) -> str:
    value = (raw_ship_date or "").strip()
    if not value:
        return ""

    normalized = value.replace("T", " ")
    for sep in (" ", "."):
        if sep in normalized:
            date_candidate = normalized.split(sep, 1)[0].strip()
            if date_candidate and any(char.isdigit() for char in date_candidate):
                normalized = date_candidate
                break

    parse_formats = (
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%m-%d-%y",
    )
    for parse_format in parse_formats:
        try:
            parsed = datetime.strptime(normalized, parse_format)
            return parsed.strftime("%m/%d/%Y")
        except ValueError:
            continue

    try:
        parsed_iso = datetime.fromisoformat(value.replace("Z", ""))
        return parsed_iso.strftime("%m/%d/%Y")
    except ValueError:
        return normalized


def _set_text_node_value(node, value: str) -> None:
    if "\n" not in value:
        node.text = value
        return

    run = node.getparent()
    insert_at = run.index(node)
    run.remove(node)

    parts = value.split("\n")
    for index, part in enumerate(parts):
        if index > 0:
            br = OxmlElement("w:br")
            run.insert(insert_at, br)
            insert_at += 1

        text_node = OxmlElement("w:t")
        if part.startswith(" ") or part.endswith(" "):
            text_node.set(qn("xml:space"), "preserve")
        text_node.text = part
        run.insert(insert_at, text_node)
        insert_at += 1


def _replace_text_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    text_nodes = paragraph._p.findall(".//w:t", paragraph._p.nsmap)
    instr_nodes = paragraph._p.findall(".//w:instrText", paragraph._p.nsmap)
    for node in [*text_nodes, *instr_nodes]:
        text = node.text or ""
        updated = text
        for source, target in replacements.items():
            if source in updated:
                updated = updated.replace(source, target)
        if updated != text:
            _set_text_node_value(node, updated)


def _replace_text_in_document(
    doc: Document, replacements: dict[str, str], *, include_xml_tree: bool = True
) -> None:
    def _replace_in_table_collection(tables: list[Table]) -> None:
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        _replace_text_in_paragraph(paragraph, replacements)
                    _replace_in_table_collection(cell.tables)

    def _replace_in_element_tree(element) -> None:
        text_nodes = element.findall(".//w:t", element.nsmap)
        instr_nodes = element.findall(".//w:instrText", element.nsmap)
        for node in [*text_nodes, *instr_nodes]:
            text = node.text or ""
            updated = text
            for source, target in replacements.items():
                if source in updated:
                    updated = updated.replace(source, target)
            if updated != text:
                _set_text_node_value(node, updated)

    for paragraph in doc.paragraphs:
        _replace_text_in_paragraph(paragraph, replacements)

    _replace_in_table_collection(doc.tables)
    if include_xml_tree:
        _replace_in_element_tree(doc.element)

    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            _replace_text_in_paragraph(paragraph, replacements)
        _replace_in_table_collection(section.header.tables)
        if include_xml_tree:
            _replace_in_element_tree(section.header._element)

        for paragraph in section.footer.paragraphs:
            _replace_text_in_paragraph(paragraph, replacements)
        _replace_in_table_collection(section.footer.tables)
        if include_xml_tree:
            _replace_in_element_tree(section.footer._element)


def _set_row_height(row, twips: int, *, exact: bool = True) -> None:
    row.height = Twips(twips)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY if exact else WD_ROW_HEIGHT_RULE.AT_LEAST


def _compact_row_text(row, font_points: float = 7.5) -> None:
    for cell in row.cells:
        cell.vertical_alignment = None
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 0.86
            for run in paragraph.runs:
                run.font.size = Pt(font_points)


def _tighten_multistop_template_rows(doc: Document) -> None:
    for table in doc.tables:
        rows = list(table.rows)
        for index, row in enumerate(rows):
            row_text = " ".join(cell.text for cell in row.cells)

            if any(token in row_text for token in ("DELIVERY_1_DC", "DELIVERY_2_DC", "DELIVERY_3_DC")):
                _set_row_height(row, 285)
                _compact_row_text(row, 7.5)
            elif any(
                token in row_text
                for token in (
                    "DELIVERY_1_ADDRESS",
                    "DELIVERY_2_ADDRESS",
                    "DELIVERY_3_ADDRESS",
                )
            ):
                _set_row_height(row, 365)
                _compact_row_text(row, 7.0)
            elif any(token in row_text for token in ("DC_1", "DC_2", "DC_3")):
                _set_row_height(row, 330)
                _compact_row_text(row, 7.0)

            if 27 <= index <= 32 and not row_text.strip():
                _set_row_height(row, 115)
                _compact_row_text(row, 7.0)


def _resolve_comment_for_record(record_comment: str, batch_comment: str | None) -> str:
    record_value = (record_comment or "").strip()
    if record_value:
        return record_value
    return (batch_comment or "").strip()


def _postprocess_comments_in_saved_docx(destination: Path, resolved_comment: str) -> bool:
    xml_path = "word/document.xml"
    with zipfile.ZipFile(destination, "r") as archive:
        if xml_path not in archive.namelist():
            return False
        file_payloads = {name: archive.read(name) for name in archive.namelist()}

    xml_text = file_payloads[xml_path].decode("utf-8", errors="ignore")
    updated_xml = xml_text
    safe_comment = escape(resolved_comment)

    if resolved_comment:
        updated_xml = updated_xml.replace("Comments:</w:t>", f"Comments: {safe_comment}</w:t>", 1)
        updated_xml = updated_xml.replace("COMMENTS:</w:t>", f"COMMENTS: {safe_comment}</w:t>", 1)

    comment_label_populated = bool(
        resolved_comment
        and (
            f"Comments: {safe_comment}</w:t>" in updated_xml
            or f"COMMENTS: {safe_comment}</w:t>" in updated_xml
        )
    )

    if updated_xml == xml_text:
        return comment_label_populated

    file_payloads[xml_path] = updated_xml.encode("utf-8")
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in file_payloads.items():
            archive.writestr(name, payload)

    return comment_label_populated


def _populate_ship_from_block(doc: Document, selected_facility: BolFacilityRecord) -> bool:
    name_value = selected_facility["facility_name"]
    street_value = selected_facility["address"]
    city_state_zip_value = selected_facility["location"]

    for table in doc.tables:
        in_ship_from_block = False
        for row in table.rows:
            row_cells = row.cells
            if not row_cells:
                continue

            row_text_upper = " ".join(cell.text.strip() for cell in row_cells).upper()
            first_cell_text = row_cells[0].text.strip().upper().replace(" ", "")

            if "FROM (SHIPPER)" in row_text_upper:
                in_ship_from_block = True
                continue
            if in_ship_from_block and "TO (CONSIGNEE)" in row_text_upper:
                return True
            if not in_ship_from_block:
                continue

            if first_cell_text == "NAME" and len(row_cells) > 1:
                row_cells[1].text = name_value
            elif first_cell_text == "STREET" and len(row_cells) > 1:
                row_cells[1].text = street_value
            elif first_cell_text in {"CITY/ST/ZIP", "CITY/STATE/ZIP"} and len(row_cells) > 1:
                row_cells[1].text = city_state_zip_value

    return False


def _template_replacements(record: BolMultistopRecord) -> dict[str, str]:
    replacements = {
        _tok("BOL_"): record.bol_number,
        _tok("ship_date"): _format_ship_date_for_template(record.ship_date),
        _tok("Carrier"): record.carrier,
        _tok("load"): record.load_number,
        _tok("KK_PO"): record.kk_po_number,
        _tok("KK_Load"): record.kk_load_number,
        _tok("DELIVERY_1_DC"): record.delivery_1_dc,
        _tok("DELIVERY_1_ADDRESS"): record.delivery_1_address,
        _tok("DELIVERY_2_DC"): record.delivery_2_dc,
        _tok("DELIVERY_2_ADDRESS"): record.delivery_2_address,
        _tok("DELIVERY_3_DC"): record.delivery_3_dc,
        _tok("DELIVERY_3_ADDRESS"): record.delivery_3_address,
        _tok("DC_1"): record.dc_1,
        _tok("CASE_1"): record.case_1,
        _tok("PO_1"): record.po_1,
        _tok("Pallet_Description_1"): record.pallet_description_1,
        _tok("PLT_1"): record.plt_1,
        _tok("WEIGHT_1"): record.weight_1,
        _tok("DC_2"): record.dc_2,
        _tok("CASE_2"): record.case_2,
        _tok("PO_2"): record.po_2,
        _tok("Pallet_Description_2"): record.pallet_description_2,
        _tok("PLT_2"): record.plt_2,
        _tok("WEIGHT_2"): record.weight_2,
        _tok("DC_3"): record.dc_3,
        _tok("CASE_3"): record.case_3,
        _tok("PO_3"): record.po_3,
        _tok("Pallet_Description_3"): record.pallet_description_3,
        _tok("PLT_3"): record.plt_3,
        _tok("WEIGHT_3"): record.weight_3,
        _tok("Total_Case"): _format_number(record.total_case),
        _tok("Total_Pallet"): _format_number(record.total_pallet),
        _tok("Total_Ship_Weight"): _format_number(record.total_ship_weight),
    }
    return replacements


def _build_individual_stop_standard_record(
    record: BolMultistopRecord,
    stop_index: int,
) -> BolStandardRecord:
    stop = record.stops[stop_index]
    return BolStandardRecord(
        bol_number=record.bol_number,
        ship_date=record.ship_date,
        carrier=record.carrier,
        kk_load_number=record.kk_load_number,
        kk_po_number=record.kk_po_number,
        po_number=stop.target_po_number,
        dc_number=stop.dc_number,
        consignee_company=stop.delivery_dc,
        consignee_street=stop.delivery_address,
        consignee_city_state_zip=stop.delivery_city_state_zip,
        ship_from=record.ship_from,
        bill_to=record.bill_to,
        seal_number_blank="",
        comments=record.comments,
        item_lines=[
            BolStandardItemLine(
                source_row_number=stop.source_row_number,
                pallet_qty=stop.cases,
                type="PLT",
                po_number=stop.target_po_number,
                item_description=stop.pallet_description,
                item_number="",
                upc="",
                skids=stop.total_pallets,
                weight_each=stop.weight,
            )
        ],
        total_skids=_parse_number(stop.cases),
        is_ready=True,
        status="Ready",
        selected_for_generation=True,
    )


def _save_multistop_docx(
    *,
    record: BolMultistopRecord,
    bol_label: str,
    selected_facility: BolFacilityRecord,
    batch_comment: str | None,
    resolved_template: Path,
    output_root: Path,
    base_name: str,
    replacements: dict[str, str],
    document_type: str,
    stop_number: int | None,
    notices: list[DocxGenerationNotice],
) -> MultistopGeneratedDocxFile:
    doc = Document(str(resolved_template))
    _tighten_multistop_template_rows(doc)
    _replace_text_in_document(doc, replacements, include_xml_tree=True)
    ship_from_populated = _populate_ship_from_block(doc, selected_facility)
    if not ship_from_populated:
        notices.append(
            DocxGenerationNotice(
                bol_number=bol_label,
                message="Could not confirm ship-from block location in template.",
            )
        )

    destination = _unique_destination_path(output_root, base_name, ".docx")
    filename = destination.name
    doc.save(str(destination))

    resolved_comment = _resolve_comment_for_record(record.comments, batch_comment)
    comment_label_populated = _postprocess_comments_in_saved_docx(destination, resolved_comment)
    if resolved_comment and not comment_label_populated:
        notices.append(
            DocxGenerationNotice(
                bol_number=bol_label,
                message=(
                    "Resolved comment was non-empty but could not be confirmed "
                    "at the visible Comments label in word/document.xml."
                ),
            )
        )

    return MultistopGeneratedDocxFile(
        bol_number=bol_label,
        file_name=filename,
        file_path=str(destination.resolve()),
        document_type=document_type,
        load_number=record.load_number,
        stop_number=stop_number,
    )


def _save_individual_stop_docx(
    *,
    record: BolMultistopRecord,
    bol_label: str,
    selected_facility: BolFacilityRecord,
    batch_comment: str | None,
    resolved_template: Path,
    output_root: Path,
    base_name: str,
    stop_index: int,
    notices: list[DocxGenerationNotice],
) -> MultistopGeneratedDocxFile:
    stop = record.stops[stop_index]
    stop_record = _build_individual_stop_standard_record(record, stop_index)
    doc = Document(str(resolved_template))
    is_standard_template = resolved_template.name == STANDARD_TEMPLATE_PATH.name
    resolved_comment = _resolve_comment_for_record(record.comments, batch_comment)
    record_notice_messages = _apply_standard_template_record_values(
        doc,
        stop_record,
        selected_facility,
        batch_comment,
        compact_standard_item_area=is_standard_template,
    )
    for message in record_notice_messages:
        notices.append(DocxGenerationNotice(bol_number=bol_label, message=message))

    destination = _unique_destination_path(output_root, base_name, ".docx")
    filename = destination.name
    doc.save(str(destination))

    comment_label_populated = _postprocess_standard_comments_in_saved_docx(
        destination,
        resolved_comment,
    )
    if resolved_comment and not comment_label_populated:
        notices.append(
            DocxGenerationNotice(
                bol_number=bol_label,
                message=(
                    "Resolved comment was non-empty but could not be confirmed "
                    "at the visible Comments label in word/document.xml."
                ),
            )
        )

    return MultistopGeneratedDocxFile(
        bol_number=bol_label,
        file_name=filename,
        file_path=str(destination.resolve()),
        document_type="stop",
        load_number=record.load_number,
        stop_number=stop.stop_number,
    )


def generate_multistop_docx_set(
    records: list[BolMultistopRecord],
    selected_facility: BolFacilityRecord | None,
    batch_comment: str | None = None,
    template_path: Path | None = None,
    individual_stop_template_path: Path | None = None,
    output_dir: Path | None = None,
    file_name_prefix: str = "multistop_bol",
) -> StandardDocxGenerationResult:
    if selected_facility is None:
        raise ValueError(
            "No ship-from facility is selected. Select a facility in BOL Generator before DOCX generation."
        )

    resolved_template = template_path or MULTISTOP_TEMPLATE_PATH
    if not resolved_template.exists():
        raise FileNotFoundError(f"Template file not found: {resolved_template}")

    resolved_individual_stop_template = individual_stop_template_path or STANDARD_TEMPLATE_PATH
    if not resolved_individual_stop_template.exists():
        raise FileNotFoundError(
            f"Individual stop template file not found: {resolved_individual_stop_template}"
        )

    output_root = output_dir or Path(mkdtemp(prefix="kkg_multistop_bol_docx_"))
    output_root.mkdir(parents=True, exist_ok=True)

    generated: list[GeneratedDocxFile] = []
    skipped: list[SkippedDocxRecord] = []
    failed: list[FailedDocxRecord] = []
    notices: list[DocxGenerationNotice] = []

    for record in records:
        bol_label = record.bol_number or "(missing BOL #)"
        record.generation_skip_reason = None

        if not record.selected_for_generation:
            reason = "Record excluded in review."
            record.generation_skip_reason = reason
            skipped.append(SkippedDocxRecord(bol_number=bol_label, reason=reason))
            continue

        if record.stop_count > 3:
            reason = "Unsupported stop count: more than 3 stops."
            record.generation_skip_reason = reason
            skipped.append(SkippedDocxRecord(bol_number=bol_label, reason=reason))
            continue

        if not record.is_ready:
            reason = "Record is not ready for DOCX generation."
            if record.status == "Unsupported Stop Count":
                reason = "Unsupported stop count: more than 3 stops."
            elif record.missing_required_fields:
                reason = "Missing required data: " + ", ".join(record.missing_required_fields)
            elif record.issues:
                reason = "; ".join(record.issues)
            record.generation_skip_reason = reason
            skipped.append(SkippedDocxRecord(bol_number=bol_label, reason=reason))
            continue

        try:
            safe_bol = _sanitize_filename_part(record.bol_number)
            safe_load = _sanitize_filename_part(record.load_number)
            combined_base_name = f"combined_multistop_bol_{safe_bol}_{safe_load}"

            generated.append(
                _save_multistop_docx(
                    record=record,
                    bol_label=bol_label,
                    selected_facility=selected_facility,
                    batch_comment=batch_comment,
                    resolved_template=resolved_template,
                    output_root=output_root,
                    base_name=combined_base_name,
                    replacements=_template_replacements(record),
                    document_type="combined",
                    stop_number=None,
                    notices=notices,
                )
            )

            for stop_index, stop in enumerate(record.stops):
                stop_base_name = f"stop_{stop.stop_number}_bol_{safe_bol}_{safe_load}"
                generated.append(
                    _save_individual_stop_docx(
                        record=record,
                        bol_label=bol_label,
                        selected_facility=selected_facility,
                        batch_comment=batch_comment,
                        resolved_template=resolved_individual_stop_template,
                        output_root=output_root,
                        base_name=stop_base_name,
                        stop_index=stop_index,
                        notices=notices,
                    )
                )
        except Exception as exc:
            failed.append(FailedDocxRecord(bol_number=bol_label, error=str(exc)))

    if not generated and not failed:
        raise ValueError("No selected and ready records are available for DOCX generation.")

    return StandardDocxGenerationResult(
        output_dir=str(output_root.resolve()),
        generated_files=generated,
        skipped_records=skipped,
        failed_records=failed,
        notices=notices,
    )
