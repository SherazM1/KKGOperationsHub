import io
import unittest
from collections import Counter

import pandas as pd

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_truck_assigner import assign_to_trucks
from app.services.truck_inventory_export import export_required_columns_excel
from app.services.truck_inventory_manual import assignment_rows, apply_manual_assignments
from app.services.truck_inventory_item_setup import build_default_item_setup, apply_dimension_group_stacking, apply_item_setup


class CrossRowStackTests(unittest.TestCase):
    def rows(self, count=3, **changes):
        return [TruckInventoryRecord(**(dict(source_type='combined', source_file='test',
            po_number=f'PO-{i}', retailer_po_number=f'PO-{i}', store_number=f'STORE-{i}',
            dc_number='6006', kkg_load_number='LOAD', item_number=f'ITEM-{i}', qty=1,
            item_length=48, item_width=48, item_height=37, item_weight=100 + i * 10,
            is_stackable=True, stack_qty=2, stack_by_dimension=True) | changes))
            for i in range(count)]

    def plan(self, records):
        return assign_to_trucks(records, preset_key='straight_box_truck_26_foot', grouping_rule='combined')

    def test_matching_single_quantity_rows_actually_stack(self):
        records = self.rows()
        trucks = self.plan(records)
        self.assertEqual(len(trucks), 1)
        truck = trucks[0]
        self.assertEqual(sorted(b.represented_qty for b in truck.boxes), [1, 2])
        self.assertEqual(truck.total_weight, 330)
        self.assertEqual(truck.total_used_pallets, 2 * 48 * 48)
        self.assertEqual(truck.pallet_positions_used, 2)
        self.assertEqual(truck.validation_status, 'valid')
        actual = Counter((m.item_number, m.store_number, m.retailer_po_number)
                         for b in truck.boxes for m in b.allocations)
        self.assertEqual(actual, Counter((r.item_number, r.store_number, r.po_number) for r in records))
        stacked = next(b for b in truck.boxes if b.represented_qty == 2)
        self.assertEqual([m.stack_level for m in stacked.allocations], [1, 2])
        self.assertEqual(stacked.grouped_height, 74)
        self.assertEqual(len({m.color for m in stacked.allocations}), 2)

    def test_exports_and_manual_moves_preserve_both_stack_members(self):
        records = self.rows(2)
        trucks = self.plan(records)
        data = io.BytesIO(export_required_columns_excel(records, trucks))
        export = pd.read_excel(data, sheet_name='Truck Export')
        self.assertEqual(export['Qty'].sum(), 2)
        self.assertEqual(set(export['Item #']), {r.item_number for r in records})
        assignments = pd.read_excel(data, sheet_name='Truck Assignments')
        self.assertEqual(set(assignments['Stack Level']), {1, 2})
        self.assertTrue(assignments['Double Stacked'].all())
        edits = assignment_rows(trucks)
        self.assertEqual(len(edits), 1)
        edits[0]['Truck #'] = 2
        moved = apply_manual_assignments(trucks, edits)
        self.assertEqual(moved[0].truck_id, 2)
        self.assertEqual(len(moved[0].boxes), 1)
        self.assertEqual({m.item_number for m in moved[0].boxes[0].allocations}, {'ITEM-0', 'ITEM-1'})
        self.assertEqual(moved[0].total_weight, 210)

    def test_unselected_dimensions_and_roof_clearance_do_not_stack(self):
        for changes in ({'stack_by_dimension': False}, {'item_height': 40.1}, {'item_height': 39.9}):
            records = self.rows(2)
            for key, value in changes.items():
                setattr(records[1], key, value)
            self.assertEqual(len(self.plan(records)[0].boxes), 2)
        tall = self.plan(self.rows(2, item_height=40, truck_height=82))
        self.assertTrue(all(b.represented_qty == 1 for t in tall for b in t.boxes))
        exact = self.plan(self.rows(2, item_height=40, truck_height=83))
        self.assertEqual(len(exact[0].boxes), 1)

    def test_weight_splitting_preserves_all_rows(self):
        records = self.rows(8, truck_max_weight=250)
        trucks = self.plan(records)
        self.assertTrue(all(t.validation_status == 'valid' and t.total_weight <= 250 for t in trucks))
        actual = Counter(m.item_number for t in trucks for b in t.boxes for m in b.allocations)
        self.assertEqual(actual, Counter(r.item_number for r in records))

    def test_rotated_half_pallet_group_stacks_with_clearance_and_moves(self):
        records = self.rows(2, pallet_type='Half', item_length=48, item_width=24, item_height=37.938)
        setup = build_default_item_setup(records)
        setup, skipped = apply_dimension_group_stacking(setup, records, [(48, 24, 37.938)], True)
        self.assertFalse(skipped)
        records, issues = apply_item_setup(records, setup)
        self.assertFalse(issues)
        truck = self.plan(records)[0]
        self.assertEqual(truck.validation_status, 'valid')
        self.assertEqual(len(truck.boxes), 1)
        self.assertEqual(truck.boxes[0].represented_qty, 2)
        self.assertEqual((truck.boxes[0].item_length, truck.boxes[0].item_width), (24, 48))
        self.assertGreaterEqual(truck.truck_height - truck.boxes[0].grouped_height, 3)
        self.assertEqual(truck.pallet_positions_used, 0.5)
        edits = assignment_rows([truck])
        edits[0]['Truck #'] = 2
        moved = apply_manual_assignments([truck], edits)
        self.assertEqual(moved[0].boxes[0].represented_qty, 2)
        tall = self.plan(self.rows(2, pallet_type='Half', item_height=41))
        self.assertTrue(all(box.represented_qty == 1 for t in tall for box in t.boxes))
