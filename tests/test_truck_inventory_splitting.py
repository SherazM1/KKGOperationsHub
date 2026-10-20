import unittest
from collections import Counter
from io import BytesIO

import pandas as pd

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_truck_assigner import assign_to_trucks
from app.services.truck_inventory_export import export_required_columns_excel
from app.utils.truck_presets import TRUCK_PRESETS


class SplitTruckTests(unittest.TestCase):
    def row(self, qty=9, **changes):
        values = dict(source_type='combined', source_file='test', po_number='PO-1',
                      kkg_load_number='LOAD-1', item_number='ITEM-1', store_number='STORE-1',
                      dc_number='DC-1', qty=qty, item_length=40, item_width=48,
                      item_height=20, item_weight=100, truck_length=80, truck_width=96,
                      truck_height=100, truck_max_weight=10000)
        values.update(changes)
        return TruckInventoryRecord(**values)

    def plan(self, rows):
        return assign_to_trucks(rows, preset_key=next(iter(TRUCK_PRESETS)), grouping_rule='combined')

    def test_space_overflow_splits_without_losing_quantity(self):
        rows = [self.row()]
        trucks = self.plan(rows)
        self.assertEqual([t.items_count for t in trucks], [4, 4, 1])
        self.assertTrue(all(t.validation_status == 'valid' for t in trucks))
        self.assertEqual(rows[0].qty, 9)
        self.assertEqual(sum(b.represented_qty for t in trucks for b in t.boxes), 9)
        self.assertTrue(all(b.store_number == 'STORE-1' for t in trucks for b in t.boxes))
        data = pd.read_excel(BytesIO(export_required_columns_excel(rows, trucks)), sheet_name='Truck Assignments')
        self.assertEqual(data['Qty'].sum(), 9)
        self.assertEqual(data['Truck #'].nunique(), 3)

    def test_weight_overflow(self):
        trucks = self.plan([self.row(qty=5, truck_max_weight=250)])
        self.assertEqual([t.items_count for t in trucks], [2, 2, 1])
        self.assertTrue(all(t.total_weight <= 250 for t in trucks))

    def test_export_load_suffix_matches_item_truck_and_quantity(self):
        rows = [self.row(qty=5, kkg_load_number='DC 6006'),
                self.row(qty=1, kkg_load_number='DC 6006', store_number='STORE-2')]
        trucks = self.plan(rows)
        data = pd.read_excel(BytesIO(export_required_columns_excel(rows, trucks)))
        actual = data.groupby(['KKG Load #', 'Store #'])['Qty'].sum().to_dict()
        expected = Counter()
        for truck in trucks:
            for box in truck.boxes:
                expected[(f'DC 6006-{truck.truck_id}', box.store_number)] += box.represented_qty
        self.assertEqual(actual, dict(expected))
        self.assertEqual(data['Qty'].sum(), 6)
        self.assertEqual(rows[0].kkg_load_number, 'DC 6006')

    def test_export_duplicate_rows_preserves_assigned_quantities(self):
        rows = [self.row(qty=3), self.row(qty=3)]
        trucks = self.plan(rows)
        data = pd.read_excel(BytesIO(export_required_columns_excel(rows, trucks)))
        self.assertEqual(data.groupby('KKG Load #')['Qty'].sum().to_dict(),
                         {f'LOAD-1-{t.truck_id}': t.items_count for t in trucks})

    def test_export_failed_item_retains_unsuffixed_load(self):
        rows = [self.row(qty=2, item_length=500), self.row(qty=1, item_number='GOOD')]
        trucks = self.plan(rows)
        data = pd.read_excel(BytesIO(export_required_columns_excel(rows, trucks)))
        self.assertEqual(data.loc[data['Item #'] == 'ITEM-1', 'KKG Load #'].tolist(), ['LOAD-1'])
        self.assertEqual(data['Qty'].sum(), 3)

    def test_backfills_earlier_trucks_to_reduce_truck_count(self):
        rows = [self.row(qty=2, item_number='HEAVY', item_weight=600, truck_max_weight=1000),
                self.row(qty=2, item_number='LIGHT', item_weight=400, truck_max_weight=1000)]
        trucks = self.plan(rows)
        self.assertEqual(len(trucks), 2)
        self.assertEqual([truck.total_weight for truck in trucks], [1000, 1000])
        self.assertTrue(all(truck.validation_status == 'valid' for truck in trucks))

    def test_double_stack_requires_six_inches_roof_clearance(self):
        exact = self.plan([self.row(qty=2, item_height=47, is_stackable=True, stack_qty=2)])
        self.assertEqual(exact[0].validation_status, 'valid')
        self.assertEqual(len(exact[0].boxes), 1)
        self.assertEqual(exact[0].boxes[0].represented_qty, 2)
        too_tall = self.plan([self.row(qty=2, item_height=47.1, is_stackable=True, stack_qty=2)])
        self.assertEqual(too_tall[0].validation_status, 'error')
        self.assertFalse(too_tall[0].boxes)
        self.assertTrue(any('6 in roof clearance' in note for note in too_tall[0].validation_notes))
        single = self.plan([self.row(qty=1, item_height=100)])
        self.assertEqual(single[0].validation_status, 'valid')

    def test_half_pallets_cannot_be_stacked(self):
        for changes in ({'pallet_type': 'Half'}, {'pallet_type': 'Half Pallet'},
                        {'item_number': '683054545'}):
            with self.subTest(changes=changes):
                trucks = self.plan([self.row(qty=2, is_stackable=True, stack_qty=2, **changes)])
                self.assertEqual(trucks[0].validation_status, 'error')
                self.assertFalse(trucks[0].boxes)
                self.assertTrue(any('half pallet' in note for note in trucks[0].validation_notes))

    def test_double_stacks_and_mixed_loads(self):
        rows = [self.row(qty=7, is_stackable=True, stack_qty=2),
                self.row(qty=5, kkg_load_number='LOAD-2', item_number='ITEM-2')]
        trucks = self.plan(rows)
        actual = Counter()
        for truck in trucks:
            self.assertEqual(truck.validation_status, 'valid')
            for box in truck.boxes:
                actual[(box.kkg_load_number, box.item_number)] += box.represented_qty
        self.assertEqual(actual, {('LOAD-1', 'ITEM-1'): 7, ('LOAD-2', 'ITEM-2'): 5})
        self.assertEqual(sum(t.total_weight for t in trucks), 1200)

    def test_impossible_pallet_is_reported_and_other_items_ship(self):
        trucks = self.plan([self.row(qty=2, item_length=500), self.row(qty=3, item_number='GOOD')])
        failed = [t for t in trucks if t.validation_status == 'error']
        self.assertEqual(len(failed), 1)
        self.assertEqual(failed[0].items_count, 2)
        self.assertEqual(sum(t.items_count for t in trucks if t.validation_status == 'valid'), 3)
