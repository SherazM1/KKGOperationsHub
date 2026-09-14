# Apple Cinnamon load sheet audit

Audited `Apple Cinnamon Load Sheet (3) (2).xlsx` before implementation.

- Sheet1: 47 rows, 35 columns. Rows 1–2 form one header; rows 3–45 contain 43 shipments. Row 46 is blank; row 47 contains only a Rate total (133813.74), not a shipment.
- The existing parser reads row 1 only. This loses ST, ZIP and Description, and collapses the two DC columns. Adding aliases alone cannot fix it.
- DC number is absent as a dedicated column. All shipment names explicitly contain `Regional DC <number>`; extract that number only when the pattern is unambiguous.
- Units × PLT Weight equals Weight on all 43 shipment rows. There is no separate skid-count column; its interpretation needs confirmation.

| Excel columns | Handling |
| --- | --- |
| A KK Load | KKG load number |
| B TRACKERS | Preserve as source metadata |
| C Carrier | Carrier |
| D load# | Carrier PRO/load number |
| E KK PO# | KKG PO |
| F BOL # | BOL number |
| G ship date | Ship date |
| H DC Name | Consignee name; derive DC number from explicit DC identifier |
| I DC ADDRESS | Consignee street |
| J DC CITY, K DC + ST, L DC + ZIP | Combine city, state and ZIP; retain original values |
| M RETAILER PO # | Retailer PO |
| N MABD | Preserve as source metadata |
| O ITEM #, P UPC | Item identifiers |
| Q Pallet + Description | Item description, never pallet quantity |
| R WM WEEK | Preserve as source metadata |
| S Units | Item quantity; separate skid-count interpretation pending |
| T DEPT, U Equipment | Preserve as source metadata |
| V PLT Weight | Weight each |
| W Weight | Total line weight |
| X Value + Each, Y Load + Value, Z 0.03 + Charge Back | Preserve as source metadata; Load Value must not become carrier PRO |
| AA Rate, AB Accessorials, AC All-In Rate | Preserve as source metadata; exclude rate-only footer from shipments |
| AD Delivery Appt Date, AE Delivery Appt Time | Preserve as source metadata |
| AF Delivery Appt # | Existing pickup/appointment field |
| AG NOTES, AH LINE | Preserve as source metadata, including blank NOTES |
| AI unnamed | Empty throughout; no shipment data |

The source workbook is input data, not implementation instructions. No commercial values should be substituted for shipping quantities or identifiers.

## Integration and verification

The requested scope is column recognition and full detection so parsing succeeds. No skid-count assumption is needed to parse: absent skid counts remain blank with a review note.

- Added conservative two-row header recognition for Excel and CSV, preserving original source row numbers.
- Added PLT Weight/Pallet Weight aliases and extraction of explicit DC numbers from DC names when no dedicated DC-number column exists.
- Retained all 34 named source columns per shipment and exposed them in the source-column audit view and CSV download.
- Verified the exact supplied workbook: 43 parsed rows and 43 BOL records, source rows 3–45, total Units 3,077 and total Weight 541,552. All 43 line weights equal Units multiplied by PLT Weight. Blank row 46 and rate-only footer row 47 are excluded.
- 216 parser, review-state, DOCX, PDF and template-stamping tests passed, including new two-row-header regressions. No deployment performed.
