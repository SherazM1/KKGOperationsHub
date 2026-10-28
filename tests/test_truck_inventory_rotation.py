import unittest
from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_item_setup import build_default_item_setup, apply_item_setup
from app.services.truck_inventory_truck_assigner import assign_to_trucks


class HalfPalletRotationTests(unittest.TestCase):
    def row(self, pallet_type='Half', length=48, width=24):
        return TruckInventoryRecord(source_type='combined', source_file='test.xlsx',
            po_number='PO', kkg_load_number='LOAD', item_number='IMPORTED-HALF', qty=8,
            pallet_type=pallet_type, item_length=length, item_width=width,
            item_height=37.938, item_weight=200)

    def test_imported_half_is_rotated_through_the_layout(self):
        records = [self.row()]
        setup = build_default_item_setup(records)
        records, issues = apply_item_setup(records, setup)
        self.assertFalse(issues)
        trucks = assign_to_trucks(records, preset_key='straight_box_truck_26_foot')
        self.assertTrue(all(truck.validation_status == 'valid' for truck in trucks))
        self.assertEqual(sum(truck.items_count for truck in trucks), 8)
        for truck in trucks:
            for box in truck.boxes:
                self.assertAlmostEqual(box.width * truck.truck_length, 24)
                self.assertAlmostEqual(box.height * truck.truck_width, 48)
                self.assertEqual(box.item_height, 37.938)
                self.assertEqual(box.item_weight, 200)

    def test_rebuilding_and_recreating_setup_do_not_flip_it_back(self):
        records = [self.row()]
        setup = build_default_item_setup(records)
        for _ in range(3):
            records, issues = apply_item_setup(records, setup)
            self.assertFalse(issues)
            self.assertEqual((records[0].item_length, records[0].item_width), (24, 48))
            setup = build_default_item_setup(records)
            self.assertEqual((setup[0]['Length'], setup[0]['Width']), (48, 24))

    def test_other_input_orientation_and_full_pallets(self):
        for kind, length, width, expected in [('Half', 20, 48, (48, 20)),
                                             ('Full', 40, 48, (40, 48))]:
            records = [self.row(kind, length, width)]
            apply_item_setup(records, build_default_item_setup(records))
            self.assertEqual((records[0].item_length, records[0].item_width), expected)
