"""Facility configuration for Standard-family BOL ship-from selection."""

from __future__ import annotations

from typing import TypedDict


class BolFacilityRecord(TypedDict):
    facility: str
    facility_name: str
    location: str
    address: str


BOL_FACILITY_RECORDS: tuple[BolFacilityRecord, ...] = (
    {
        "facility": "PRODUCTIV-PATRIOT",
        "facility_name": "Kendal King C/O Productiv",
        "location": "Grapevine,TX",
        "address": "4255 Patriot Drive Suite 400 Grapevine, TX",
    },
    {
        "facility": "SHORR-LANCASTER",
        "facility_name": "Kendal King C/O Shorr",
        "location": "Lancaster,TX",
        "address": "2900 West Drive Lancaster, TX 75134",
    },
    {
        "facility": "MIDWEST FULFILLMENT",
        "facility_name": "Kendal King C/O Midwest Fulfillment",
        "location": "Berkley, MO",
        "address": "9083 Frost Ave Berkley, MO 63134",
    },
    {
        "facility": "SHORR",
        "facility_name": "Kendal King C/O Shorr",
        "location": "Grand Prairie, TX",
        "address": "975 W Oakdale Road Grand Prairie, TX 75050",
    },
    {
        "facility": "WE PACK-107",
        "facility_name": "Kendal King C/O We Pack",
        "location": "Paris, TX",
        "address": "2300 SW 13th Street Paris, TX 75460",
    },
    {
        "facility": "WAREHOUSE PRO",
        "facility_name": "Kendal King C/O Warehouse Pro",
        "location": "Rockwall,TX",
        "address": "2020 Industrial Blvd Rockwall, TX 75087",
    },
    {
        "facility": "STRIBLING-LOWELL",
        "facility_name": "Kendal King C/O Stribling",
        "location": "Lowell, AR",
        "address": "419 South Lincoln Street Suite A Lowell, AR 72745",
    },
    {
        "facility": "PRODUCTIV-ESTERS",
        "facility_name": "Kendal King C/O Productiv",
        "location": "Grapevine,TX",
        "address": "2450 Esters BLVD Suite 100 Grapevine, TX 76051",
    },
    {
        "facility": "STRIBLING-ROGERS",
        "facility_name": "Kendal King C/O Stribling",
        "location": "Rogers, AR",
        "address": "1603 N 35th Street, Rogers, AR 72756",
    },
    {
        "facility": "LOGIC WAREHOUSE",
        "facility_name": "Kendal King C/O Logic",
        "location": "Kansas City, MO",
        "address": "1329 Quebec Street, Kansas City, MO 64116",
    },
    {
        "facility": "CTL GLOBAL",
        "facility_name": "Kendal King C/O CTL Global",
        "location": "Elmhurst, IL",
        "address": "1000 N County Line Road Elmhurst, IL 60126",
    },
    {
        "facility": "KENDAL KING LAB",
        "facility_name": "Kendal King Lab",
        "location": "Bentonville, AR",
        "address": "901 SW A St. Bentonville, AR 72712",
    },
    {
        "facility": "MIDAMERICA",
        "facility_name": "Kendal King C/O Midamerica",
        "location": "Bridgeton, MO",
        "address": "111 Boulder Industrial Drive Bridgeton, MO 63044",
    },
    {
        "facility": "KINTER",
        "facility_name": "Kendal King C/O Kinter",
        "location": "Waukegan, IL",
        "address": "3333 Oak Grove Ave Waukegan, IL 60087",
    },
    {
        "facility": "RAND GRAPHICS",
        "facility_name": "Kendal King C/O Rand Graphics",
        "location": "Wichita, KS",
        "address": "2820 S. Hoover Road Wichita, KS 67215",
    },
    {
        "facility": "SMC PACKAGING - ARROWHEAD",
        "facility_name": "Kendal King C/O SMC Packaging",
        "location": "Kansas City, MO",
        "address": "4330 Clary Blvd. Kansas City, MO 64130",
    },
    {
        "facility": "SHORR-WEST CHICAGO",
        "facility_name": "Kendal King C/O Shorr",
        "location": "West Chicago, IL",
        "address": "555 Innovation Drive West Chicago, IL 60185",
    },
    {
        "facility": "LAMB AND ASSOCIATES PACKAGING",
        "facility_name": "Kendal King C/O Lamb and Associates",
        "location": "Maumelle, AR",
        "address": "1700 Murphy Drive Maumelle, AR 72113",
    },
    {
        "facility": "Titan Corrugated",
        "facility_name": "Kendal King C/O Titan Corrugated",
        "location": "Flower Mound, TX",
        "address": "801 Lakeside PWKY Flower Mound, TX 75028",
    },
    {
        "facility": "JA Warehousing",
        "facility_name": "Kendal King C/O JA Warehousing",
        "location": "St. Louis, MO",
        "address": "10750 Baur BLVD St. Louis, MO 63132",
    },
    {
        "facility": "Green Bay Packaging",
        "facility_name": "Kendal King C/O Green Bay",
        "location": "New Berlin, WI",
        "address": "5600 S. Moorland Road New Berlin, WI 53151",
    },
)

BOL_FACILITY_OPTIONS: tuple[str, ...] = tuple(
    facility_record["facility"] for facility_record in BOL_FACILITY_RECORDS
)

BOL_FACILITY_LOOKUP: dict[str, BolFacilityRecord] = {
    facility_record["facility"]: facility_record for facility_record in BOL_FACILITY_RECORDS
}
