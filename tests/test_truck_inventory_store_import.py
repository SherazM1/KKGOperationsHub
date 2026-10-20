"""Store inventory workbook import and physical setup regression tests."""
import io
import unittest

import pandas as pd

from app.services.truck_inventory_parser import parse_combined_load_sheet
from app.services.truck_inventory_normalizer import normalize_rows
from app.services.truck_inventory_item_setup import build_default_item_setup, apply_item_setup
from app.services.truck_inventory_validator import validate_records
from app.services.truck_inventory_export import export_required_columns_excel


class StoreImportTests(unittest.TestCase):
    def parse(self, rows):
        data = io.BytesIO()
        pd.DataFrame(rows).to_excel(data, index=False)
        data.name = 'store_inventory.xlsx'
        data.seek(0)
        result = parse_combined_load_sheet(data)
        self.assertTrue(result.success, result.error_message)
        return normalize_rows(result.rows, data.name, 'combined_load_sheet')

    def row(self, **overrides):
        return {'STORE #': '00477', 'DC #': '06006', 'ITEM #': '662022666',
                'QUANTITY': 1, 'PALLET TYPE': 'Full', 'PALLET LENGTH': 48,
                'PALLET WIDTH': 48, 'PALLET HEIGHT': 70.688, 'WEIGHT': 350,
                **overrides}

    def test_store_identifiers_dimensions_setup_and_export(self):
        records = self.parse([self.row(), self.row(**{'STORE #': '00656'})])
        record = records[0]
        self.assertEqual(record.store_number, '00477')
        self.assertEqual(record.dc_number, '06006')
        self.assertEqual(record.kkg_load_number, 'DC-06006')
        self.assertEqual(record.po_number, '')
        self.assertEqual(record.item_height, 70.688)
        self.assertEqual(validate_records(records)[1].errors, 0)
        setup = build_default_item_setup(records)
        self.assertEqual([setup[0][key] for key in ['Length', 'Width', 'Height', 'Weight']], [48, 48, 70.688, 350])
        _, issues = apply_item_setup(records, setup)
        self.assertFalse(issues)
        exported = pd.read_excel(io.BytesIO(export_required_columns_excel(records)), dtype=object)
        self.assertEqual(exported.iloc[0]['Store #'], '00477')
        self.assertEqual(exported.iloc[0]['Pallet Height'], 70.688)

    def test_conflicting_item_dimensions_require_review(self):
        records = self.parse([self.row(), self.row(**{'PALLET HEIGHT': 40})])
        setup = build_default_item_setup(records)
        self.assertEqual(setup[0]['Height'], 0)
        self.assertTrue(apply_item_setup(records, setup)[1])

    def test_missing_store_is_reported(self):
        records = self.parse([self.row(**{'STORE #': ''})])
        self.assertEqual(validate_records(records)[1].errors, 1)
        self.assertIn('Missing Store #', records[0].validation_notes)

    def test_fractional_quantity_is_not_silently_rounded_down(self):
        records = self.parse([self.row(**{'QUANTITY': 1.5})])
        self.assertIsNone(records[0].qty)
        self.assertEqual(validate_records(records)[1].errors, 1)

    def test_original_load_sheet_format_still_imports(self):
        records = self.parse([{'KKG Load #': 'LOAD-1', 'Retailer PO #': 'PO-1', 'Item #': 'ITEM-1', 'Qty': 2}])
        self.assertEqual(records[0].kkg_load_number, 'LOAD-1')
        self.assertEqual(records[0].retailer_po_number, 'PO-1')
        self.assertEqual(validate_records(records)[1].errors, 0)
