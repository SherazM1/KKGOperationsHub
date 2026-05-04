"""Grouping and mapping service for Multistop BOL records."""

from __future__ import annotations

from collections import defaultdict

from app.models.bol_multistop_record import BolMultistopRecord, BolMultistopStop
from app.models.bol_multistop_row import BolMultistopRow
from app.models.bol_standard_record import BolAddressBlock, BolStandardItemLine


MAX_SUPPORTED_STOPS = 3


def _first_non_empty(rows: list[BolMultistopRow], attr_name: str) -> str:
    for row in rows:
        value = getattr(row, attr_name, "")
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _parse_number(value: str) -> float | None:
    cleaned = (value or "").replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def _build_delivery_dc(dc_name: str, dc_number: str) -> str:
    name = (dc_name or "").strip()
    number = (dc_number or "").strip()
    if name and number:
        return f"{name} ({number})"
    return name or number


def _build_city_state_zip(row: BolMultistopRow) -> str:
    combined = (row.dc_city_state_zip or "").strip()
    if combined:
        return combined

    city = (row.dc_city or "").strip()
    state = (row.dc_state or "").strip()
    zip_code = (row.dc_zip or "").strip()

    city_state = ", ".join(part for part in [city, state] if part)
    return " ".join(part for part in [city_state, zip_code] if part).strip()


def _build_delivery_address(street: str, city_state_zip: str) -> str:
    return "\n".join(
        part.strip() for part in [street, city_state_zip] if (part or "").strip()
    )


def _sum_numeric(
    rows: list[BolMultistopRow],
    attr_name: str,
    label: str,
    issues: list[str],
) -> tuple[float, bool]:
    total = 0.0
    parse_failed = False
    has_numeric_value = False

    for row in rows:
        raw_value = getattr(row, attr_name, "")
        if not (raw_value or "").strip():
            issues.append(f"Row {row.source_row_number}: missing {label} input.")
            continue

        parsed = _parse_number(raw_value)
        if parsed is None:
            parse_failed = True
            issues.append(
                f"Row {row.source_row_number}: {label} is not numeric ('{raw_value}')."
            )
            continue

        has_numeric_value = True
        total += parsed

    if parse_failed:
        issues.append(f"{label} total could not be computed reliably.")

    return total, has_numeric_value


def _header_consistency_warnings(group_rows: list[BolMultistopRow]) -> list[str]:
    warnings: list[str] = []
    fields = [
        ("Ship Date", "ship_date"),
        ("Carrier", "carrier"),
        ("Load #", "load_number"),
        ("KKG PO#", "kk_po_number"),
        ("KKG Load #", "kk_load"),
    ]

    for label, attr in fields:
        values = sorted(
            {
                getattr(row, attr).strip()
                for row in group_rows
                if isinstance(getattr(row, attr), str) and getattr(row, attr).strip()
            }
        )
        if len(values) > 1:
            warnings.append(
                f"Inconsistent {label} values in group: {', '.join(values[:3])}"
                + (" ..." if len(values) > 3 else "")
            )

    return warnings


def _build_stop(row: BolMultistopRow, stop_number: int) -> BolMultistopStop:
    return BolMultistopStop(
        source_row_number=row.source_row_number,
        stop_number=stop_number,
        delivery_dc=_build_delivery_dc(row.dc_name, row.dc_number),
        delivery_address=row.dc_address,
        delivery_city_state_zip=_build_city_state_zip(row),
        dc_number=row.dc_number,
        cases=row.cases,
        target_po_number=row.target_po_number,
        pallet_description=row.pallet_description,
        item_number=row.item_number,
        upc=row.upc,
        total_pallets=row.total_pallets,
        weight=row.weight,
    )


def _validate_stop_fields(stop: BolMultistopStop) -> list[str]:
    issues: list[str] = []
    required_fields = [
        ("delivery DC", stop.delivery_dc),
        ("delivery address", stop.delivery_address),
        ("DC #", stop.dc_number),
        ("Cases", stop.cases),
        ("TGT PO #", stop.target_po_number),
        ("PalletDescription", stop.pallet_description),
        ("Total PLT", stop.total_pallets),
        ("Weight", stop.weight),
    ]
    for field_label, value in required_fields:
        if not (value or "").strip():
            issues.append(
                f"Stop {stop.stop_number} (row {stop.source_row_number}): missing {field_label}."
            )
    return issues


def _is_duplicate_stop_warning(warning: str) -> bool:
    return warning.startswith("Duplicate Stop ")


def _build_item_lines(stops: list[BolMultistopStop], rows_by_stop: dict[int, BolMultistopRow]) -> list[BolStandardItemLine]:
    item_lines: list[BolStandardItemLine] = []
    for stop in stops:
        source_row = rows_by_stop[stop.stop_number]
        item_lines.append(
            BolStandardItemLine(
                source_row_number=stop.source_row_number,
                pallet_qty=stop.cases,
                type="PLT",
                po_number=stop.target_po_number,
                item_description=stop.pallet_description,
                item_number=source_row.item_number,
                upc=source_row.upc,
                skids=stop.total_pallets,
                weight_each=source_row.weight_each or stop.weight,
            )
        )
    return item_lines


def _group_key(row: BolMultistopRow) -> tuple[str, str]:
    return (row.bol_number.strip(), row.load_number.strip())


def _optional_grouped_field_warnings(group_rows: list[BolMultistopRow]) -> list[str]:
    warnings: list[str] = []
    optional_fields = [
        ("Ship Date", _first_non_empty(group_rows, "ship_date")),
        ("Load #", _first_non_empty(group_rows, "load_number")),
        ("KKG PO#", _first_non_empty(group_rows, "kk_po_number")),
        ("KKG Load #", _first_non_empty(group_rows, "kk_load")),
    ]
    for field_name, value in optional_fields:
        if not (value or "").strip():
            warnings.append(f"Optional field missing: {field_name}")
    return warnings


def map_multistop_rows_to_records(rows: list[BolMultistopRow]) -> list[BolMultistopRecord]:
    grouped_rows: dict[tuple[str, str], list[BolMultistopRow]] = defaultdict(list)
    for row in rows:
        grouped_rows[_group_key(row)].append(row)

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

    records: list[BolMultistopRecord] = []

    for group, group_rows in grouped_rows.items():
        bol_number, load_number = group
        issues: list[str] = []
        consistency_warnings = _header_consistency_warnings(group_rows)
        optional_warnings = _optional_grouped_field_warnings(group_rows)
        validation_warnings: list[str] = []

        rows_sorted = sorted(
            group_rows,
            key=lambda row: (
                row.stop_number if row.stop_number is not None else 9999,
                row.source_row_number,
            ),
        )

        rows_by_stop: dict[int, BolMultistopRow] = {}
        distinct_stops: set[int] = set()
        malformed_stop_rows = 0

        for row in rows_sorted:
            if row.stop_number is None or row.stop_number <= 0:
                malformed_stop_rows += 1
                if not row.stop.strip():
                    validation_warnings.append(
                        f"Row {row.source_row_number}: missing Stop value."
                    )
                else:
                    validation_warnings.append(
                        f"Row {row.source_row_number}: malformed Stop value ('{row.stop}')."
                    )
                continue

            distinct_stops.add(row.stop_number)
            if row.stop_number in rows_by_stop:
                validation_warnings.append(
                    f"Duplicate Stop {row.stop_number} rows found; using first row in mapping."
                )
                continue
            rows_by_stop[row.stop_number] = row

        has_unsupported_stops = any(stop_number > MAX_SUPPORTED_STOPS for stop_number in distinct_stops)
        stop_count = len(distinct_stops)
        if stop_count > MAX_SUPPORTED_STOPS or has_unsupported_stops:
            issues.append(
                f"Unsupported stop count: found {stop_count} stops with max stop {max(distinct_stops) if distinct_stops else 0}; maximum supported is {MAX_SUPPORTED_STOPS}."
            )

        ordered_stops: list[BolMultistopStop] = []
        for stop_number in range(1, MAX_SUPPORTED_STOPS + 1):
            row = rows_by_stop.get(stop_number)
            if row is None:
                continue
            stop = _build_stop(row, stop_number)
            ordered_stops.append(stop)
            issues.extend(_validate_stop_fields(stop))

        if distinct_stops:
            highest_supported = min(max(distinct_stops), MAX_SUPPORTED_STOPS)
            for expected_stop in range(1, highest_supported + 1):
                if expected_stop not in rows_by_stop:
                    issues.append(f"Missing Stop {expected_stop} row in this grouped shipment.")
        elif malformed_stop_rows:
            issues.append("No valid stop numbers were found in this grouped shipment.")

        total_case, has_case_values = _sum_numeric(group_rows, "cases", "Cases", issues)
        total_pallet, has_pallet_values = _sum_numeric(group_rows, "total_pallets", "Total PLT", issues)
        total_ship_weight, has_weight_values = _sum_numeric(group_rows, "weight", "Weight", issues)

        missing_required_fields: list[str] = []
        header_required = [
            ("BOL #", bol_number),
            ("Carrier", _first_non_empty(group_rows, "carrier")),
        ]
        for field_name, value in header_required:
            if not (value or "").strip():
                missing_required_fields.append(field_name)
                issues.append(f"Missing required grouped field: {field_name}.")

        if not ordered_stops:
            missing_required_fields.append("STOP ROWS")
            issues.append("No valid stops were mapped for this grouped shipment.")
        if not has_case_values:
            missing_required_fields.append("Cases")
            issues.append("Missing totals input for Cases.")
        if not has_pallet_values:
            missing_required_fields.append("Total PLT")
            issues.append("Missing totals input for Total PLT.")
        if not has_weight_values:
            missing_required_fields.append("Weight")
            issues.append("Missing totals input for Weight.")

        for warning in validation_warnings:
            if _is_duplicate_stop_warning(warning):
                issues.append(warning)
            else:
                issues.append(warning)

        stop1_row = rows_by_stop.get(1)

        record = BolMultistopRecord(
            bol_number=bol_number,
            ship_date=_first_non_empty(group_rows, "ship_date"),
            carrier=_first_non_empty(group_rows, "carrier"),
            load_number=load_number,
            kk_po_number=_first_non_empty(group_rows, "kk_po_number"),
            kk_load_number=_first_non_empty(group_rows, "kk_load"),
            group_key=f"{bol_number}::{load_number}",
            stop_count=stop_count,
            stops=ordered_stops,
            delivery_1_dc=ordered_stops[0].delivery_dc if len(ordered_stops) > 0 else "",
            delivery_1_address=(
                _build_delivery_address(
                    ordered_stops[0].delivery_address,
                    ordered_stops[0].delivery_city_state_zip,
                )
                if len(ordered_stops) > 0
                else ""
            ),
            delivery_2_dc=ordered_stops[1].delivery_dc if len(ordered_stops) > 1 else "",
            delivery_2_address=(
                _build_delivery_address(
                    ordered_stops[1].delivery_address,
                    ordered_stops[1].delivery_city_state_zip,
                )
                if len(ordered_stops) > 1
                else ""
            ),
            delivery_3_dc=ordered_stops[2].delivery_dc if len(ordered_stops) > 2 else "",
            delivery_3_address=(
                _build_delivery_address(
                    ordered_stops[2].delivery_address,
                    ordered_stops[2].delivery_city_state_zip,
                )
                if len(ordered_stops) > 2
                else ""
            ),
            dc_1=ordered_stops[0].dc_number if len(ordered_stops) > 0 else "",
            case_1=ordered_stops[0].cases if len(ordered_stops) > 0 else "",
            po_1=ordered_stops[0].target_po_number if len(ordered_stops) > 0 else "",
            pallet_description_1=ordered_stops[0].pallet_description if len(ordered_stops) > 0 else "",
            plt_1=ordered_stops[0].total_pallets if len(ordered_stops) > 0 else "",
            weight_1=ordered_stops[0].weight if len(ordered_stops) > 0 else "",
            dc_2=ordered_stops[1].dc_number if len(ordered_stops) > 1 else "",
            case_2=ordered_stops[1].cases if len(ordered_stops) > 1 else "",
            po_2=ordered_stops[1].target_po_number if len(ordered_stops) > 1 else "",
            pallet_description_2=ordered_stops[1].pallet_description if len(ordered_stops) > 1 else "",
            plt_2=ordered_stops[1].total_pallets if len(ordered_stops) > 1 else "",
            weight_2=ordered_stops[1].weight if len(ordered_stops) > 1 else "",
            dc_3=ordered_stops[2].dc_number if len(ordered_stops) > 2 else "",
            case_3=ordered_stops[2].cases if len(ordered_stops) > 2 else "",
            po_3=ordered_stops[2].target_po_number if len(ordered_stops) > 2 else "",
            pallet_description_3=ordered_stops[2].pallet_description if len(ordered_stops) > 2 else "",
            plt_3=ordered_stops[2].total_pallets if len(ordered_stops) > 2 else "",
            weight_3=ordered_stops[2].weight if len(ordered_stops) > 2 else "",
            total_case=total_case,
            total_pallet=total_pallet,
            total_ship_weight=total_ship_weight,
            po_number=(stop1_row.target_po_number if stop1_row else _first_non_empty(group_rows, "target_po_number")),
            dc_number=(stop1_row.dc_number if stop1_row else _first_non_empty(group_rows, "dc_number")),
            consignee_company=(stop1_row.dc_name if stop1_row else _first_non_empty(group_rows, "dc_name")),
            consignee_street=(stop1_row.dc_address if stop1_row else _first_non_empty(group_rows, "dc_address")),
            consignee_city_state_zip=(
                _build_city_state_zip(stop1_row)
                if stop1_row
                else _build_city_state_zip(group_rows[0])
            ),
            ship_from=ship_from,
            bill_to=bill_to,
            seal_number_blank="",
            comments="",
            item_lines=_build_item_lines(ordered_stops, rows_by_stop),
            total_skids=total_case,
            is_ready=False,
            status="Missing Required Data",
            selected_for_generation=True,
            missing_required_fields=missing_required_fields,
            warnings=consistency_warnings + optional_warnings + validation_warnings,
            generation_skip_reason=None,
            conversion_skip_reason=None,
            issues=list(dict.fromkeys(issues + consistency_warnings + validation_warnings)),
            is_supported=not (stop_count > MAX_SUPPORTED_STOPS or has_unsupported_stops),
        )

        if stop_count > MAX_SUPPORTED_STOPS or has_unsupported_stops:
            record.is_ready = False
            record.status = "Unsupported Stop Count"
        elif missing_required_fields:
            record.is_ready = False
            record.status = "Missing Required Data"
        elif record.issues:
            record.is_ready = False
            record.status = "Warning"
        else:
            record.is_ready = True
            record.status = "Ready"

        records.append(record)

    return records
