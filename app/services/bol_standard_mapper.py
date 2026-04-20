"""Grouping and mapping service for Standard BOL records."""

from __future__ import annotations

from collections import defaultdict

from app.models.bol_standard_record import (
    BolAddressBlock,
    BolStandardItemLine,
    BolStandardRecord,
)
from app.models.bol_standard_row import BolStandardRow


def _first_non_empty(rows: list[BolStandardRow], attr_name: str) -> str:
    for row in rows:
        value = getattr(row, attr_name, "")
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _parse_number(value: str) -> float | None:
    cleaned = value.replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def _required_shipment_issues(record: BolStandardRecord) -> list[str]:
    issues: list[str] = []

    required_fields = [
        ("BOL #", record.bol_number),
        ("SHIP DATE", record.ship_date),
        ("CARRIER", record.carrier),
        ("KK LOAD", record.kk_load_number),
        ("KK PO#", record.kk_po_number),
        ("WM PO #", record.po_number),
        ("DC #", record.dc_number),
        ("DC NAME", record.consignee_company),
        ("DC STREET", record.consignee_street),
        ("DC CITY, STATE, ZIP", record.consignee_city_state_zip),
    ]

    for field_name, value in required_fields:
        if not value.strip():
            issues.append(f"Missing required shipment field: {field_name}.")

    if not record.item_lines:
        issues.append("No item lines found for this BOL.")

    for line in record.item_lines:
        if not line.item_description.strip():
            issues.append(f"Row {line.source_row_number}: missing item description.")
        if not line.item_number.strip():
            issues.append(f"Row {line.source_row_number}: missing ITEM #.")
        if not line.upc.strip():
            issues.append(f"Row {line.source_row_number}: missing UPC.")
        if not line.pallet_qty.strip():
            issues.append(f"Row {line.source_row_number}: missing QTY.")
        if not line.weight_each.strip():
            issues.append(f"Row {line.source_row_number}: missing WEIGHT EACH.")

    return issues


def _missing_required_fields(record: BolStandardRecord) -> list[str]:
    missing: list[str] = []
    required_fields = [
        ("BOL #", record.bol_number),
        ("SHIP DATE", record.ship_date),
        ("CARRIER", record.carrier),
        ("KK LOAD", record.kk_load_number),
        ("KK PO#", record.kk_po_number),
        ("WM PO #", record.po_number),
        ("DC #", record.dc_number),
        ("DC NAME", record.consignee_company),
        ("DC STREET", record.consignee_street),
        ("DC CITY, STATE, ZIP", record.consignee_city_state_zip),
    ]
    for field_name, value in required_fields:
        if not value.strip():
            missing.append(field_name)
    if not record.item_lines:
        missing.append("ITEM ROWS")
    return missing


def _inconsistent_shipment_warnings(bol_rows: list[BolStandardRow]) -> list[str]:
    warnings: list[str] = []
    fields_to_check = [
        ("SHIP DATE", "ship_date"),
        ("CARRIER", "carrier"),
        ("KK LOAD", "kk_load"),
        ("KK PO#", "kk_po"),
        ("WM PO #", "wm_po"),
        ("DC #", "dc_number"),
        ("DC NAME", "dc_name"),
        ("DC STREET", "dc_street"),
        ("DC CITY, STATE, ZIP", "dc_city_state_zip"),
    ]

    for field_label, attr_name in fields_to_check:
        values = sorted(
            {
                getattr(row, attr_name).strip()
                for row in bol_rows
                if isinstance(getattr(row, attr_name), str)
                and getattr(row, attr_name).strip()
            }
        )
        if len(values) > 1:
            warnings.append(
                f"Inconsistent {field_label} values within this BOL: {', '.join(values[:3])}"
                + (" ..." if len(values) > 3 else "")
            )

    return warnings


def map_standard_rows_to_records(rows: list[BolStandardRow]) -> list[BolStandardRecord]:
    grouped_rows: dict[str, list[BolStandardRow]] = defaultdict(list)

    for row in rows:
        grouped_rows[row.bol_number.strip()].append(row)

    ship_from = BolAddressBlock(
        company="Kendal King C/O Shorr",
        street="981 W Oakdale Rd",
        city_state_zip="Grand Prairie, TX 75056",
        attn="",
    )
    bill_to = BolAddressBlock(
        company="Trident Transport, LLC",
        street="505 Riverfront Pkwy",
        city_state_zip="Chattanooga, TN 37402",
        attn="",
    )

    records: list[BolStandardRecord] = []

    for bol_key, bol_rows in grouped_rows.items():
        item_lines: list[BolStandardItemLine] = []
        total_skids = 0.0

        for row in bol_rows:
            item_line = BolStandardItemLine(
                source_row_number=row.source_row_number,
                pallet_qty=row.quantity,
                type="PLT",
                po_number=row.wm_po,
                item_description=row.item_description,
                item_number=row.item_number,
                upc=row.upc,
                skids=row.quantity,
                weight_each=row.weight_each,
            )
            item_lines.append(item_line)

            qty_number = _parse_number(row.quantity)
            if qty_number is not None:
                total_skids += qty_number

        record = BolStandardRecord(
            bol_number=bol_key,
            ship_date=_first_non_empty(bol_rows, "ship_date"),
            carrier=_first_non_empty(bol_rows, "carrier"),
            kk_load_number=_first_non_empty(bol_rows, "kk_load"),
            kk_po_number=_first_non_empty(bol_rows, "kk_po"),
            po_number=_first_non_empty(bol_rows, "wm_po"),
            dc_number=_first_non_empty(bol_rows, "dc_number"),
            consignee_company=_first_non_empty(bol_rows, "dc_name"),
            consignee_street=_first_non_empty(bol_rows, "dc_street"),
            consignee_city_state_zip=_first_non_empty(bol_rows, "dc_city_state_zip"),
            ship_from=ship_from,
            bill_to=bill_to,
            seal_number_blank="",
            comments="",
            item_lines=item_lines,
            total_skids=total_skids,
            is_ready=True,
            status="Ready",
            selected_for_generation=True,
            missing_required_fields=[],
            warnings=[],
            generation_skip_reason=None,
            conversion_skip_reason=None,
            issues=[],
        )

        missing_required = _missing_required_fields(record)
        warnings = _inconsistent_shipment_warnings(bol_rows)
        issues = _required_shipment_issues(record)
        record.missing_required_fields = missing_required
        record.warnings = warnings
        record.issues = issues + warnings

        if missing_required:
            record.is_ready = False
            record.status = "Missing Required Data"
        elif warnings:
            record.is_ready = False
            record.status = "Warning"
        else:
            record.is_ready = True
            record.status = "Ready"

        records.append(record)

    return records
