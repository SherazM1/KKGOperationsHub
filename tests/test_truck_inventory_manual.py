import unittest
from collections import Counter

from app.services.truck_inventory_manual import assignment_rows, apply_manual_assignments
from tests import test_truck_inventory_splitting as splitting


class ManualTruckTests(unittest.TestCase):
    def plan(self):
        helper = splitting.SplitTruckTests()
        return helper.plan([helper.row(qty=5)])

    def test_move_to_new_truck_preserves_quantity_and_identifiers(self):
        original = self.plan()
        rows = assignment_rows(original)
        rows[0]['Truck #'] = 3
        updated = apply_manual_assignments(original, rows)
        self.assertEqual([t.truck_id for t in updated], [1, 2, 3])
        self.assertEqual(sum(t.items_count for t in updated), 5)
        self.assertEqual(Counter((b.item_number, b.store_number, b.dc_number, b.color)
                                 for t in original for b in t.boxes),
                         Counter((b.item_number, b.store_number, b.dc_number, b.color)
                                 for t in updated for b in t.boxes))

    def test_overloaded_edit_does_not_mutate_original(self):
        original = self.plan()
        before = assignment_rows(original)
        rows = assignment_rows(original)
        for row in rows:
            row['Truck #'] = 1
        with self.assertRaisesRegex(ValueError, 'Truck 1'):
            apply_manual_assignments(original, rows)
        self.assertEqual(assignment_rows(original), before)

    def test_cannot_drop_or_duplicate_pallets(self):
        original = self.plan()
        rows = assignment_rows(original)
        with self.assertRaises(ValueError):
            apply_manual_assignments(original, rows[:-1])
        rows[0]['Pallet ID'] = rows[1]['Pallet ID']
        with self.assertRaises(ValueError):
            apply_manual_assignments(original, rows)

    def test_order_changes_placement(self):
        helper = splitting.SplitTruckTests()
        original = helper.plan([helper.row(qty=1, store_number='FIRST'), helper.row(qty=1, store_number='SECOND')])
        rows = assignment_rows(original)
        rows[0]['Order'] = 99
        updated = apply_manual_assignments(original, rows)
        self.assertEqual(updated[0].boxes[0].store_number, 'SECOND')
        self.assertEqual(updated[0].boxes[1].store_number, 'FIRST')

    def test_double_stack_moves_as_one_unit(self):
        helper = splitting.SplitTruckTests()
        original = helper.plan([helper.row(qty=4, is_stackable=True, stack_qty=2)])
        rows = assignment_rows(original)
        rows[0]['Truck #'] = 2
        updated = apply_manual_assignments(original, rows)
        self.assertEqual([t.items_count for t in updated], [2, 2])
        self.assertTrue(all(b.represented_qty == 2 for t in updated for b in t.boxes))
