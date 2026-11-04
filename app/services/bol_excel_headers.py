"""Shared, deterministic header detection for BOL Excel and CSV imports."""

from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, time
import re
import unicodedata
import math
import csv
from io import StringIO

import pandas as pd

from openpyxl.utils import get_column_letter


HEADER_SCAN_ROWS = 20
ColumnLookups = tuple[dict[str, list[str]], dict[str, list[str]], dict[str, list[str]]]


class ColumnMappingRequired(ValueError):
    """An import needs an explicit source-column choice before it can continue."""

    def __init__(self, message, requests, columns=()):
        super().__init__(message)
        self.requests = {field: tuple(options) for field, options in requests.items()}
        self.columns = tuple(columns)
        self.samples = {}


def normalize_header(value):
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = re.sub(r"[\u200b\ufeff]", "", text)
    text = re.sub(r"\s*#\s*", "#", text)
    return re.sub(r"\s+", " ", text).strip().upper()


def original_header(value):
    return re.sub(r" \[column [A-Z]+\]$", "", str(value))


def compact_header(value):
    return re.sub(r"[^A-Z0-9#%]+", "", normalize_header(original_header(value)))


def build_column_lookups(columns):
    exact, normalized, compact = {}, {}, {}
    for column in map(str, columns):
        exact.setdefault(original_header(column), []).append(column)
        normalized.setdefault(normalize_header(original_header(column)), []).append(column)
        compact.setdefault(compact_header(column), []).append(column)
    return exact, normalized, compact


def resolve_column_name(lookups, primary, aliases=()):
    # Alias priority distinguishes item-level PLT QTY from load-level totals.
    # Columns equivalent to the SAME alias must never be chosen by position.
    for candidate in (primary, *aliases):
        options = lookups[2].get(compact_header(candidate), [])
        if len(options) > 1:
            raise ColumnMappingRequired(
                f"Ambiguous column for {primary}: {', '.join(options)}. Choose the source column.",
                {primary: options},
            )
        if options:
            return options[0]
    return None


@dataclass(frozen=True)
class HeaderLayout:
    headers: list[str]
    first_data_row: int
    header_row: int


_CONTINUATIONS = {
    ("DC", "ST"), ("DC", "STATE"), ("DC", "ZIP"),
    ("PALLET", "DESCRIPTION"), ("ITEM", "DESCRIPTION"),
    ("VALUE", "EACH"), ("LOAD", "VALUE"), ("003", "CHARGEBACK"),
    ("TOTALPLT", "QTY"), ("TOTALPLT", "QUANTITY"),
    ("TOTALLOAD", "WEIGHT"), ("TOTAL", "PLT"),
}
_IDENTITY_HEADERS = {"BOL#", "BOL", "BOLNUMBER", "BILLOFLADING", "BILLOFLADINGNUMBER",
                     "KKLOAD", "KKLOAD#", "KKGLOAD", "KKGLOAD#", "CARRIER", "ITEM#", "UPC"}


def _text(value):
    return "" if value is None or isinstance(value, float) and math.isnan(value) else str(value).strip()


def _continuation(headers, row):
    populated = [(i, value) for i, value in enumerate(row) if _text(value)]
    if not populated:
        return False
    # A shipment row can contain text such as ST or Description too. Its IDs,
    # dates or numeric quantities prevent it being consumed as a header row.
    if any(isinstance(value, (int, float, bool, date, datetime, time)) for _, value in populated):
        return False
    if any(compact_header(headers[i]) in _IDENTITY_HEADERS for i, _ in populated if i < len(headers)):
        return False
    return any((compact_header(headers[i]), compact_header(value)) in _CONTINUATIONS
               for i, value in populated if i < len(headers))


def detect_headers(rows, score):
    """Find up to three header rows, preserving the first shipment row."""
    candidates = []
    for index, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        headers = [_text(value) for value in row]
        initial_score = score(headers)
        if initial_score < 6:
            continue
        consumed = 1
        for below in rows[index + 1:index + 3]:
            if not _continuation(headers, below):
                break
            for column, value in enumerate(below):
                if _text(value):
                    headers[column] = f"{headers[column]} {_text(value)}".strip()
            consumed += 1
        candidates.append((score(headers), -index, headers, index + consumed + 1, index + 1))
    if not candidates:
        headers = [_text(value) for value in rows[0]] if rows else []
        return HeaderLayout(headers, 2, 1)
    _, _, headers, first_data_row, header_row = max(candidates, key=lambda item: item[:2])
    counts = Counter(normalize_header(header) for header in headers if header)
    headers = [f"{header} [column {get_column_letter(i + 1)}]"
               if header and counts[normalize_header(header)] > 1 else header
               for i, header in enumerate(headers)]
    return HeaderLayout(headers, first_data_row, header_row)


def apply_column_override(field, overrides, columns):
    selected = (overrides or {}).get(field)
    if selected is not None and selected not in columns:
        raise ValueError(f"The saved source column for {field} is no longer present. Choose it again.")
    return selected


def attach_mapping_samples(error, headers, rows):
    preview = list(rows)[:3]
    error.samples = {
        header: [_text(row[i]) if i < len(row) else "" for row in preview]
        for i, header in enumerate(headers) if header
    }
    return error


def read_csv_source(file):
    """Preserve duplicate headers, leading zeros and ragged title/header rows."""
    file.seek(0)
    content = file.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8-sig")
    return pd.DataFrame(list(csv.reader(StringIO(content))), dtype=object)


def is_weight_summary(values):
    """A weight-only footer is not a shipment; incomplete item rows remain visible."""
    populated = {key for key, value in values.items() if str(value or "").strip()}
    return bool(populated) and populated <= {"total_weight", "weight", "group_weight_total"}
