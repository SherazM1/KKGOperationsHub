"""Combined load fit and double-stack behavior."""
import unittest
from io import BytesIO

import pandas as pd

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_export import export_required_columns_excel
from app.services.truck_inventory_item_setup import apply_item_setup
from app.services.truck_inventory_truck_assigner import assign_to_trucks
from app.services.truck_inventory_item_setup import normalize_editor_rows as _normalize_editor_rows
from app.utils.truck_presets import TRUCK_PRESETS


class CombinedTruckTests(unittest.TestCase):
    def records(self, qty=3, height=20):
        return [TruckInventoryRecord(source_type='combined', source_file='test.xlsx',
                    po_number=f'PO-{load}', retailer_po_number=f'PO-{load}',
                    kkg_load_number=load, item_number='ITEM-1', qty=qty,
                    item_length=40, item_width=48, item_height=height, item_weight=100)
                for load in ['LOAD-A', 'LOAD-B']]

    def plan(self, records, **kwargs):
        return assign_to_trucks(records, preset_key=next(iter(TRUCK_PRESETS)), **kwargs)

    def test_combines_loads_and_preserves_source_identifiers(self):
        records = self.records()
        separate = self.plan(records)
        combined = self.plan(records, grouping_rule='combined')
        self.assertEqual(len(separate), 2)
        self.assertEqual(len(combined), 1)
        truck = combined[0]
        self.assertEqual(truck.load_groups, ['LOAD-A', 'LOAD-B'])
        self.assertEqual(truck.total_weight, 600)
        self.assertTrue(truck.weight_fit and truck.spatial_fit)
        self.assertEqual({(b.kkg_load_number, b.retailer_po_number) for b in truck.boxes},
                         {('LOAD-A', 'PO-LOAD-A'), ('LOAD-B', 'PO-LOAD-B')})
        exported = pd.read_excel(BytesIO(export_required_columns_excel(records)))
        self.assertEqual(set(exported['KKG Load #']), {'LOAD-A', 'LOAD-B'})

    def test_double_stack_odd_quantity_keeps_all_units_and_weight(self):
        records = self.records()
        single = self.plan(records, grouping_rule='combined')[0]
        setup = _normalize_editor_rows([{'Item #': 'ITEM-1', 'Length': 40, 'Width': 48,
                    'Height': 20, 'Weight': 100, 'Double Stack': True}])
        records, issues = apply_item_setup(records, setup)
        self.assertFalse(issues)
        stacked = self.plan(records, grouping_rule='combined')[0]
        self.assertEqual(len(single.boxes), 6)
        self.assertEqual(len(stacked.boxes), 4)
        self.assertEqual(sum(b.represented_qty for b in stacked.boxes), 6)
        self.assertEqual(stacked.total_weight, single.total_weight)
        setup[0] = _normalize_editor_rows([{**setup[0], 'Double Stack': False}])[0]
        records, _ = apply_item_setup(records, setup)
        self.assertEqual(len(self.plan(records, grouping_rule='combined')[0].boxes), 6)

    def test_combined_weight_and_stack_height_limits(self):
        records = self.records(height=100)
        for record in records:
            record.is_stackable = True
            record.stack_qty = 2
        truck = self.plan(records, grouping_rule='combined', operational_weight_threshold_lbs=500)[0]
        self.assertTrue(truck.weight_fit)
        self.assertFalse(truck.spatial_fit)
        self.assertEqual(truck.validation_status, 'error')


if __name__ == '__main__':
    unittest.main()
