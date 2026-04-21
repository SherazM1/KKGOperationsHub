"""DOCX generation service for Standard-family BOL records."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from tempfile import mkdtemp

from docx import Document
from docx.table import Table

from app.models.bol_standard_record import BolStandardItemLine, BolStandardRecord


STANDARD_TEMPLATE_PATH = Path("app/templates/standard_bol_template.docx")
NO_RECOURSE_TEMPLATE_PATH = Path("app/templates/no_recourse_bol_template.docx")
DEFAULT_TEMPLATE_PATH = STANDARD_TEMPLATE_PATH
LEFT_MERGE = "\u00ab"
RIGHT_MERGE = "\u00bb"
ITEM_TOKEN_ALIASES: dict[str, tuple[str, ...]] = {
    "QTY": ("QTY", "QTY_2", "QTY_3", "QTY_4"),
    "TYPE": ("TYPE", "TYPE_2", "TYPE_3", "TYPE_4"),
    "PO": ("PO_", "PO_2", "PO_3", "PO_4"),
    "ITEM_DESCRIPTION": (
        "Item_Description",
        "Item_2_Description",
        "Item_Description_3",
        "Item_4_Description",
    ),
    "ITEM_NUMBER": ("Item_Number", "Item_Number_2", "Item_Number_3", "Item_Number_4"),
    "UPC": ("UPC_", "UPC__2", "UPC__3", "UPC__4"),
    "WEIGHT": ("WEIGHT", "WEIGHT_2", "WEIGHT_3", "WEIGHT_4"),
}


def _tok(name: str) -> str:
    return f"{LEFT_MERGE}{name}{RIGHT_MERGE}"


ITEM_PLACEHOLDER_TOKENS: tuple[str, ...] = tuple(
    _tok(alias) for aliases in ITEM_TOKEN_ALIASES.values() for alias in aliases
)


def resolve_template_path_for_mode(mode: str) -> Path:
    if mode == "Standard":
        return STANDARD_TEMPLATE_PATH
    if mode == "No Recourse":
        return NO_RECOURSE_TEMPLATE_PATH
    raise ValueError(
        f"Unsupported BOL mode for Standard-family generation: {mode}. "
        "Use Standard or No Recourse."
    )


def resolve_output_filename_prefix_for_mode(mode: str) -> str:
    if mode == "Standard":
        return "standard_bol"
    if mode == "No Recourse":
        return "no_recourse_bol"
    raise ValueError(
        f"Unsupported BOL mode for Standard-family generation: {mode}. "
        "Use Standard or No Recourse."
    )


@dataclass(slots=True)
class GeneratedDocxFile:
    bol_number: str
    file_name: str
    file_path: str


@dataclass(slots=True)
class SkippedDocxRecord:
    bol_number: str
    reason: str


@dataclass(slots=True)
class FailedDocxRecord:
    bol_number: str
    error: str


@dataclass(slots=True)
class DocxGenerationNotice:
    bol_number: str
    message: str


@dataclass(slots=True)
class StandardDocxGenerationResult:
    output_dir: str
    generated_files: list[GeneratedDocxFile]
    skipped_records: list[SkippedDocxRecord]
    failed_records: list[FailedDocxRecord]
    notices: list[DocxGenerationNotice]

    @property
    def generated_count(self) -> int:
        return len(self.generated_files)

    @property
    def skipped_count(self) -> int:
        return len(self.skipped_records)

    @property
    def failed_count(self) -> int:
        return len(self.failed_records)


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


def _replace_text_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    runs = paragraph.runs
    if not runs:
        return

    combined_text = "".join(run.text for run in runs)
    updated_text = combined_text
    for source, target in replacements.items():
        if source in updated_text:
            updated_text = updated_text.replace(source, target)

    if updated_text == combined_text:
        return

    original_lengths = [len(run.text) for run in runs]
    cursor = 0
    last_idx = len(runs) - 1
    for idx, run in enumerate(runs):
        if idx == last_idx:
            run.text = updated_text[cursor:]
            break

        take = min(original_lengths[idx], max(0, len(updated_text) - cursor))
        run.text = updated_text[cursor:cursor + take]
        cursor += take


def _replace_text_in_document(doc: Document, replacements: dict[str, str]) -> None:
    for paragraph in doc.paragraphs:
        _replace_text_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_text_in_paragraph(paragraph, replacements)


def _replace_tokens_in_row_element(row_element, replacements: dict[str, str]) -> None:
    text_nodes = row_element.findall(".//w:t", row_element.nsmap)
    for node in text_nodes:
        text = node.text or ""
        for source, target in replacements.items():
            if source in text:
                text = text.replace(source, target)
        node.text = text


def _row_has_any_token_text(row, tokens: tuple[str, ...]) -> bool:
    row_text = " ".join(cell.text for cell in row.cells)
    return any(token in row_text for token in tokens)


def _document_contains_token(doc: Document, token: str) -> bool:
    return token in doc.element.xml


def _override_consignee_street(doc: Document, consignee_street: str) -> None:
    for table in doc.tables:
        in_consignee_block = False
        for row in table.rows:
            row_cells = row.cells
            if not row_cells:
                continue

            first_cell_text = row_cells[0].text.strip().upper()
            row_text = " ".join(cell.text.strip() for cell in row_cells).upper()

            if "TO (CONSIGNEE)" in row_text:
                in_consignee_block = True
                continue

            if in_consignee_block and first_cell_text == "STREET" and len(row_cells) > 1:
                row_cells[1].text = consignee_street
                return

            if in_consignee_block and "QTY" in row_text and "ITEM DESCRIPTION" in row_text:
                return


def _item_row_replacements(line: BolStandardItemLine) -> dict[str, str]:
    replacements: dict[str, str] = {}
    for alias in ITEM_TOKEN_ALIASES["QTY"]:
        replacements[_tok(alias)] = line.pallet_qty
    for alias in ITEM_TOKEN_ALIASES["TYPE"]:
        replacements[_tok(alias)] = "PLT"
    for alias in ITEM_TOKEN_ALIASES["PO"]:
        replacements[_tok(alias)] = line.po_number
    for alias in ITEM_TOKEN_ALIASES["ITEM_DESCRIPTION"]:
        replacements[_tok(alias)] = line.item_description
    for alias in ITEM_TOKEN_ALIASES["ITEM_NUMBER"]:
        replacements[_tok(alias)] = line.item_number
    for alias in ITEM_TOKEN_ALIASES["UPC"]:
        replacements[_tok(alias)] = line.upc
    for alias in ITEM_TOKEN_ALIASES["WEIGHT"]:
        replacements[_tok(alias)] = line.weight_each
    return replacements


def _total_qty_display(total_skids: float) -> str:
    return str(int(total_skids)) if float(total_skids).is_integer() else str(total_skids)


def _populate_item_table(
    table: Table,
    item_lines: list[BolStandardItemLine],
    total_skids: float,
    *,
    compact_standard_item_area: bool = False,
) -> None:
    header_idx = None
    item_row_indices: list[int] = []
    total_qty_idx = None
    totals_label_idx = None

    for idx, row in enumerate(table.rows):
        row_text = " ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
        row_text_upper = row_text.upper()

        if (
            header_idx is None
            and "QTY" in row_text_upper
            and "TYPE" in row_text_upper
            and "PO #" in row_text_upper
            and "ITEM DESCRIPTION" in row_text_upper
        ):
            header_idx = idx
            continue

        if header_idx is not None and idx > header_idx and _row_has_any_token_text(
            row, ITEM_PLACEHOLDER_TOKENS
        ):
            item_row_indices.append(idx)

        if header_idx is not None and idx > header_idx and _tok("TOTAL_QTY") in row_text:
            total_qty_idx = idx
        if header_idx is not None and idx > header_idx and "TOTALS" in row_text_upper:
            if totals_label_idx is None:
                totals_label_idx = idx

    if header_idx is None:
        raise ValueError("Could not locate the item-table header in the DOCX template.")
    if not item_row_indices:
        raise ValueError("Could not locate item template rows in the DOCX template.")
    if total_qty_idx is None:
        raise ValueError("Could not locate the TOTAL_QTY row in the DOCX template.")

    first_idx = item_row_indices[0]
    contiguous_item_row_indices = [first_idx]
    for idx in item_row_indices[1:]:
        if idx == contiguous_item_row_indices[-1] + 1:
            contiguous_item_row_indices.append(idx)
        else:
            break

    insertion_anchor_idx = total_qty_idx
    if compact_standard_item_area and totals_label_idx is not None:
        insertion_anchor_idx = totals_label_idx

    table_xml = table._tbl
    template_trs = [deepcopy(table.rows[idx]._tr) for idx in contiguous_item_row_indices]
    anchor_tr = table.rows[insertion_anchor_idx]._tr

    if compact_standard_item_area:
        remove_start_idx = contiguous_item_row_indices[0]
        remove_end_idx = insertion_anchor_idx - 1
        rows_to_remove = list(range(remove_start_idx, remove_end_idx + 1))
    else:
        rows_to_remove = contiguous_item_row_indices

    for idx in sorted(rows_to_remove, reverse=True):
        table_xml.remove(table.rows[idx]._tr)

    for idx, line in enumerate(item_lines):
        new_tr = deepcopy(template_trs[idx % len(template_trs)])
        _replace_tokens_in_row_element(new_tr, _item_row_replacements(line))
        anchor_tr.addprevious(new_tr)

    totals_replacements = {
        _tok("TOTAL_QTY"): _total_qty_display(total_skids),
        _tok("TOTAL_WEIGHT"): "",
    }
    for row in table.rows:
        _replace_tokens_in_row_element(row._tr, totals_replacements)


def _apply_template_record_values(
    doc: Document,
    record: BolStandardRecord,
    *,
    compact_standard_item_area: bool = False,
) -> list[str]:
    notices: list[str] = []
    comments_value = record.comments.strip()
    has_comments_placeholder = _document_contains_token(doc, _tok("COMMENTS"))
    if comments_value and not has_comments_placeholder:
        notices.append(
            "Template has no <<COMMENTS>> placeholder; comments were not inserted for this record."
        )

    replacements = {
        _tok("BOL"): record.bol_number,
        _tok("SHIP_DATE"): record.ship_date,
        _tok("CARRIER"): record.carrier,
        _tok("Carrier_Pro_"): "",
        _tok("HOST_PO"): record.po_number,
        _tok("KKG_PO"): record.kk_po_number,
        _tok("KKG_LOAD_"): record.kk_load_number,
        _tok("Pick_Up_"): "",
        _tok("TRACKER_"): "",
        _tok("COMMENTS"): comments_value if has_comments_placeholder else "",
        _tok("SHIP_FROM"): record.ship_from.company,
        _tok("SHIP_FROM_ADDRESS"): record.ship_from.street,
        _tok("SHIP_FROM_CITY_STATE_ZIP"): record.ship_from.city_state_zip,
        _tok("SHIP_TO_NAME"): record.consignee_company,
        _tok("SHIP_TO_ADDRESS"): record.consignee_street,
        _tok("SHIP_TO_CITY_STATE_ZIP"): record.consignee_city_state_zip,
        _tok("DC"): record.dc_number,
        _tok("BILL_TO"): record.bill_to.company,
        _tok("BILL_TO_ADDRESS"): record.bill_to.street,
        _tok("BILL_TO_CITY_SATE_ZIP"): record.bill_to.city_state_zip,
    }
    _replace_text_in_document(doc, replacements)
    _override_consignee_street(doc, record.consignee_street)

    last_error: Exception | None = None
    for table in doc.tables:
        try:
            _populate_item_table(
                table,
                record.item_lines,
                record.total_skids,
                compact_standard_item_area=compact_standard_item_area,
            )
            return notices
        except ValueError as exc:
            last_error = exc

    if last_error is None:
        raise ValueError("Could not locate the item table in the DOCX template.")
    raise ValueError(f"Could not populate item table: {last_error}")


def generate_standard_docx_set(
    records: list[BolStandardRecord],
    template_path: Path | None = None,
    output_dir: Path | None = None,
    file_name_prefix: str = "standard_bol",
) -> StandardDocxGenerationResult:
    resolved_template = template_path or DEFAULT_TEMPLATE_PATH
    if not resolved_template.exists():
        raise FileNotFoundError(f"Template file not found: {resolved_template}")

    output_root = output_dir or Path(mkdtemp(prefix="kkg_standard_bol_docx_"))
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
        if not record.is_ready:
            reason = "Record is not ready for DOCX generation."
            if record.missing_required_fields:
                reason = "Missing required data: " + ", ".join(record.missing_required_fields)
            elif record.issues:
                reason = "; ".join(record.issues)
            record.generation_skip_reason = reason
            skipped.append(SkippedDocxRecord(bol_number=bol_label, reason=reason))
            continue

        try:
            doc = Document(str(resolved_template))
            is_standard_template = resolved_template.name == STANDARD_TEMPLATE_PATH.name
            record_notices = _apply_template_record_values(
                doc,
                record,
                compact_standard_item_area=is_standard_template,
            )

            safe_bol = _sanitize_filename_part(record.bol_number)
            destination = _unique_destination_path(
                output_root, f"{file_name_prefix}_{safe_bol}", ".docx"
            )
            filename = destination.name
            doc.save(str(destination))

            generated.append(
                GeneratedDocxFile(
                    bol_number=bol_label,
                    file_name=filename,
                    file_path=str(destination.resolve()),
                )
            )
            for message in record_notices:
                notices.append(DocxGenerationNotice(bol_number=bol_label, message=message))
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
