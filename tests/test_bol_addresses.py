from dataclasses import replace

import pytest

from app.utils.bol_addresses import normalize_bol_address
from tests.test_bol_standard_pdf_generator import _ready_record
from tests.test_bol_multistop_grouping import _row
from app.services.bol_multistop_mapper import map_multistop_rows_to_records


@pytest.mark.parametrize("company", ["SAM'S DC# 8234", "SAM'S DC# 8234 (8234)"])
def test_shifted_facility_and_full_address(company):
    assert normalize_bol_address(
        company, "SAMS DC# 8234", "3301 EAST PARK & BLASS AVE, SEARCY AR"
    ) == ("3301 EAST PARK & BLASS AVE", "SEARCY AR")


def test_duplicate_name_in_multiline_street():
    assert normalize_bol_address(
        "Sam's DC 8234", "SAMS DC 8234\n3301 East Park & Blass Ave", "Searcy, AR 72143"
    ) == ("3301 East Park & Blass Ave", "Searcy, AR 72143")


def test_correct_address_is_preserved():
    assert normalize_bol_address("Test DC", "123 Test Street Suite 10", "Dallas, TX 75001") == (
        "123 Test Street Suite 10", "Dallas, TX 75001"
    )


def test_missing_street_does_not_invent_address():
    assert normalize_bol_address("Test DC", "Test DC", "Dallas, TX 75001") == ("", "Dallas, TX 75001")


def test_shared_standard_record_corrects_address_before_generation():
    record = replace(_ready_record(), consignee_company="SAMS DC# 8234",
                     consignee_street="SAMS DC# 8234",
                     consignee_city_state_zip="3301 EAST PARK & BLASS AVE, SEARCY AR")
    assert record.consignee_street == "3301 EAST PARK & BLASS AVE"
    assert record.consignee_city_state_zip == "SEARCY AR"


def test_every_multistop_delivery_and_compatibility_address_is_corrected():
    rows = [replace(_row(kk_load="1", stop=n, bol_number=str(n)),
                    dc_name="SAMS DC# 8234", dc_number="8234",
                    dc_address="SAMS DC# 8234",
                    dc_city_state_zip="3301 EAST PARK & BLASS AVE, SEARCY AR")
            for n in (1, 2, 3)]
    record = map_multistop_rows_to_records(rows)[0]
    for stop in record.stops:
        assert stop.delivery_address == "3301 EAST PARK & BLASS AVE"
        assert stop.delivery_city_state_zip == "SEARCY AR"
    for n in (1, 2, 3):
        assert getattr(record, f"delivery_{n}_address") == "3301 EAST PARK & BLASS AVE\nSEARCY AR"
    assert record.consignee_street == "3301 EAST PARK & BLASS AVE"
