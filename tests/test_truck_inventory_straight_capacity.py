import unittest

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_truck_assigner import assign_to_trucks
from app.services.truck_inventory_manual import assignment_rows, apply_manual_assignments


class StraightCapacityTests(unittest.TestCase):
    def row(self, length, qty):
        return TruckInventoryRecord(source_type='test', source_file='', po_number='PO',
            kkg_load_number='LOAD', item_number=str(length), qty=qty, item_length=length,
            item_width=48, item_height=40, item_weight=100)

    def plan(self, rows):
        return assign_to_trucks(rows, preset_key='straight_box_truck_26_foot', grouping_rule='combined')

    def test_full_pallet_limit(self):
        self.assertEqual([t.items_count for t in self.plan([self.row(48, 11)])], [10, 1])
        self.assertEqual(self.plan([self.row(48, 10)])[0].load_capacity_display, '100%')

    def test_half_pallet_limit(self):
        self.assertEqual([t.items_count for t in self.plan([self.row(20, 21)])], [20, 1])
        self.assertEqual(self.plan([self.row(20, 20)])[0].load_capacity_display, '100%')

    def test_ten_full_presets_cannot_accept_extra_halves(self):
        for full_length in (40, 48):
            with self.subTest(full_length=full_length):
                rows = [self.row(full_length, 10), self.row(20, 2)]
                trucks = self.plan(rows)
                self.assertEqual(len(trucks), 2)
                self.assertEqual(trucks[0].items_count, 10)
                self.assertEqual(trucks[0].load_capacity_display, '100%')
                self.assertEqual(trucks[1].items_count, 2)
                edits = assignment_rows(trucks)
                for row in edits:
                    row['Truck #'] = 1
                with self.assertRaisesRegex(ValueError, '10 pallet positions'):
                    apply_manual_assignments(trucks, edits)

    def test_mixed_positions_and_manual_edits(self):
        trucks = self.plan([self.row(48, 6), self.row(20, 10)])
        self.assertEqual(len(trucks), 2)
        self.assertEqual(trucks[0].items_count, 14)
        edits = assignment_rows(trucks)
        for row in edits:
            row['Truck #'] = 1
        with self.assertRaisesRegex(ValueError, '10 pallet positions'):
            apply_manual_assignments(trucks, edits)
