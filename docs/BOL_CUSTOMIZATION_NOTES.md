# BOL customization code notes

## Broker names and billing addresses

- `app/utils/bol_brokers.py` stores the eight brokers from `Bill To Addresses.xlsx`, with blank strings for address details that were not supplied.
- `app/ui/bol_generator.py` initializes the broker dropdown and manual name fields, and clears generated downloads when their values change.
- `_records_with_broker_name()` copies each record and its billing block so selecting a broker cannot overwrite imported records or the shared library.
- A library selection replaces the full billing block, including blank address lines; manual name entry changes only the company name and retains the original address.
- The manual Bill to street/PO box and city/state/ZIP fields replace both billing address lines when either field is supplied, clear the old attention contact, and retain the original address when both fields are blank.
- Manual address fields are disabled for library selections, and changing them clears generated downloads so the next Word and PDF files use the latest details.
- The copied records flow to both Word and PDF generation, including Multistop masters and individual stops.
- `_populate_broker_of_record()` in `bol_standard_docx_generator.py` replaces the fixed no-recourse broker name even when Word splits it across text runs.
- `bol_pdf_template_stamper.py` removes the fixed template broker text and prints the selected Bill to company in the no-recourse notice.

## Facility and From Company controls

- `_generation_facility()` selects either the dropdown facility or the complete manual facility details and then applies any optional From Company suffix.
- Manual facility fields are stored in session state, require a name, street, city/state, and ZIP code, and apply to the current batch without adding a permanent dropdown entry.
- The From Company suffix prints after `Kendal King C/O` and preserves the facility address.
- The manual facility preview shows the final company and address, including any From Company override.
- `_apply_selected_facility_to_grouped_records()` refreshes the review records, and all generation actions receive the same resolved facility.
- Generation buttons stay disabled while a selected manual facility is incomplete.

## Multistop consignee lines

- `_clear_multistop_consignee_rules()` covers the measured internal PDF writing rules before drawing addresses, with a small margin to remove rendering fringes.
- The helper is called for Standard Multistop masters and Multistop individual stop PDFs only; standalone Standard and No Recourse output retains its rules.
- No Recourse Multistop masters already clear the consignee interior, so their overlay now draws delivery labels and addresses without redrawing rules.
- `_remove_multistop_consignee_rules()` disables the top and bottom borders of Word consignee value cells while preserving the heading, label cells, and section outline.
- The Word helper runs only in the Multistop generator, for masters and individual stops.
- If PDF templates change, recheck the measured rule positions and render both Multistop options before updating these coordinates.

## Verification

- `tests/test_bol_generator_parse_state.py` covers manual facility completeness, address construction, and returning to the original facility.
- `tests/test_bol_pdf_template_stamper.py` checks manual broker names, library billing addresses, blank addresses, and consistent master/stop Word and PDF output.
- Existing Standard and Multistop generation tests check that the shared changes preserve the other document fields and output bundles.
