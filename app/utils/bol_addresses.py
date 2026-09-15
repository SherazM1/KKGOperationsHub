"""Keep facility names out of BOL address lines."""

import re


def normalize_bol_address(company: str, street: str, city_state_zip: str) -> tuple[str, str]:
    """Remove a repeated company and recover a shifted street without guessing a city."""
    def key(value: str) -> str:
        return re.sub(r"[^a-z0-9]", "", value.lower())

    names = {key(company), key(re.sub(r"\s*\(\d+\)\s*$", "", company))} - {""}
    street_lines = [line.strip() for line in street.splitlines() if line.strip()]
    while street_lines and key(street_lines[0]) in names:
        street_lines.pop(0)
    street = "\n".join(street_lines)
    city_state_zip = city_state_zip.strip()
    if not street and re.match(r"^(?:\d+\s|P\.?\s*O\.?\s+BOX\b)", city_state_zip, re.I):
        street, city_state_zip = city_state_zip, ""
    # Split only at a recognizable street suffix followed by a city and state.
    # This also handles a full address stored on a single street line.
    if street and not city_state_zip:
        lines = street.splitlines()
        if len(lines) > 1 and re.search(r"\b[A-Z]{2}(?:\s+\d{5}(?:-\d{4})?)?$", lines[-1], re.I):
            return "\n".join(lines[:-1]), lines[-1]
        match = re.match(
            r"^(.+?\b(?:ROAD|RD|STREET|ST|AVENUE|AVE|DRIVE|DR|BOULEVARD|BLVD|"
            r"LANE|LN|COURT|CT|PARKWAY|PKWY|HIGHWAY|HWY|WAY|PLACE|PL)\.?)"
            r"[,\s]+([A-Za-z][A-Za-z .'-]*?,?\s+[A-Z]{2}(?:\s+\d{5}(?:-\d{4})?)?)$",
            street, re.I,
        )
        if match:
            return match.group(1), match.group(2)
    return street, city_state_zip
