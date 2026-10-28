import unittest

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_truck_assigner import assign_to_trucks
from app.services.truck_inventory_manual import assignment_rows, apply_manual_assignments
from app.utils.truck_presets import TRUCK_PRESETS


class TruckClearanceTests(unittest.TestCase):
    def row(self, height, qty=1, stacked=False, item='ITEM'):
        return TruckInventoryRecord(source_type='test', source_file='', po_number='PO',
            kkg_load_number='LOAD', item_number=item, qty=qty, item_length=48,
            item_width=48, item_height=height, item_weight=100,
            is_stackable=stacked, stack_qty=2 if stacked else 1)

    def test_single_and_double_stack_boundaries_for_every_truck(self):
        for key, preset in TRUCK_PRESETS.items():
            required = 6 if key == 'reefer_48_foot' else 3
            for qty in (1, 2):
                with self.subTest(truck=key, qty=qty):
                    height = (preset.height_in - required) / qty
                    truck = assign_to_trucks([self.row(height, qty, qty == 2)], preset_key=key)[0]
                    self.assertEqual(truck.roof_clearance_in, required)
                    self.assertEqual(truck.validation_status, 'valid')
                    self.assertAlmostEqual(preset.height_in - truck.boxes[0].grouped_height, required)
                    failed = assign_to_trucks([self.row(height + .01, qty, qty == 2)], preset_key=key)[0]
                    self.assertEqual(failed.validation_status, 'error')
                    self.assertTrue(any(f'{required} in roof clearance' in note for note in failed.validation_notes))

    def test_dimension_group_can_stack_with_three_inches_but_not_less(self):
        rows = [self.row(40.5, item=str(i)) for i in range(2)]
        for row in rows:
            row.is_stackable, row.stack_qty, row.stack_by_dimension = True, 2, True
        truck = assign_to_trucks(rows, preset_key='straight_box_truck_26_foot')[0]
        self.assertEqual(len(truck.boxes), 1)
        self.assertEqual(truck.boxes[0].represented_qty, 2)
        self.assertEqual(truck.roof_clearance_in, 3)

    def test_manual_move_uses_destination_truck_clearance(self):
        dry = assign_to_trucks([self.row(91, item='TALL')], preset_key='dry_van_53_foot')[0]
        reefer = assign_to_trucks([self.row(40, item='SHORT')], preset_key='reefer_48_foot')[0]
        dry.truck_id, reefer.truck_id = 1, 2
        edits = assignment_rows([dry, reefer])
        edits[0]['Truck #'] = 2
        with self.assertRaisesRegex(ValueError, '6 in roof clearance'):
            apply_manual_assignments([dry, reefer], edits)
        edits[0]['Truck #'] = 1
        moved = apply_manual_assignments([dry, reefer], edits)
        self.assertEqual([truck.roof_clearance_in for truck in moved], [3, 6])
