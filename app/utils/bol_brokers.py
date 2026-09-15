"""Store the broker billing addresses supplied in Bill To Addresses.xlsx."""

from app.models.bol_standard_record import BolAddressBlock


# These entries preserve the supplied address details, with missing address lines left blank.
BOL_BROKER_LOOKUP: dict[str, BolAddressBlock] = {
    "Trident Transport": BolAddressBlock("Trident Transport", "", "Chattanooga, TN 37402"),
    "Southland Transportation": BolAddressBlock("Southland Transportation", "PO Box 99", "Boonville, NC"),
    "TQL": BolAddressBlock("TQL", "PO Box 1160", "Smithville, MO 64089"),
    "Rite Way Logistics": BolAddressBlock("Rite Way Logistics", "", ""),
    "Arrive Logistics": BolAddressBlock("Arrive Logistics", "7701 Metropolis Dr Bldg 15", "Austin, TX 78744"),
    "Axle Logistics": BolAddressBlock("Axle Logistics", "835 N Central Street", "Knoxville, TN 37917"),
    "JB Hunt": BolAddressBlock("JB Hunt", "PO BOX 682", "Lowell, AR 72745"),
    "KLLM": BolAddressBlock("KLLM", "135 Riverview Drive", "Jackson, MS 39218"),
}

BOL_BROKER_MANUAL_OPTION = "Existing billing details / manual broker name"
