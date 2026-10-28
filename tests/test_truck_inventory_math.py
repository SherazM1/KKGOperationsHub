"""Independent arithmetic and geometry checks for generated truck plans."""
import random
import unittest
from collections import Counter

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_truck_assigner import assign_to_trucks
from app.utils.truck_presets import TRUCK_PRESETS


class TruckMathTests(unittest.TestCase):
    def row(self, **changes):
        data = dict(source_type='test', source_file='', po_number='PO',
                    kkg_load_number='LOAD', item_number='ITEM', qty=1,
                    item_length=40, item_width=48, item_height=40, item_weight=100,
                    truck_length=120, truck_width=96, truck_height=100, truck_max_weight=1000)
        return TruckInventoryRecord(**(data | changes))

    def plan(self, rows):
        return assign_to_trucks(rows, preset_key=next(iter(TRUCK_PRESETS)), grouping_rule='combined')

    def test_decimal_exact_fit(self):
        trucks = self.plan([self.row(qty=3, item_length=0.1, truck_length=0.3,
                                     item_weight=0.1, truck_max_weight=0.3)])
        self.assertEqual(len(trucks), 1)
        self.assertAlmostEqual(trucks[0].total_weight, 0.3)

    def test_invalid_quantity_and_infinite_dimensions(self):
        for overrides in ({'qty': 0}, {'qty': -1}, {'qty': 1.5}, {'item_length': float('inf')},
                          {'item_height': float('nan')}, {'is_stackable': True, 'stack_qty': 0}):
            with self.subTest(overrides=overrides):
                truck = self.plan([self.row(**overrides)])[0]
                self.assertEqual(truck.validation_status, 'error')
                self.assertEqual(truck.boxes, [])

    def test_single_pallet_leftover_does_not_require_double_height(self):
        truck = self.plan([self.row(qty=1, item_height=97, is_stackable=True, stack_qty=2)])[0]
        self.assertEqual(truck.validation_status, 'valid')
        self.assertEqual(truck.boxes[0].represented_qty, 1)

    def test_generated_plans_conserve_quantities_weight_and_geometry(self):
        rng = random.Random(28)
        for scenario in range(30):
            rows = [self.row(item_number=str(i), qty=rng.randint(1, 14),
                             item_length=rng.choice([20, 30, 40, 48]),
                             item_width=rng.choice([24, 40, 48]),
                             item_weight=rng.choice([50, 125, 250]),
                             is_stackable=bool(i % 2), stack_qty=2 if i % 2 else 1)
                    for i in range(4)]
            trucks = self.plan(rows)
            actual = Counter()
            for truck in trucks:
                self.assertEqual(truck.validation_status, 'valid')
                self.assertLessEqual(truck.total_weight, truck.truck_max_weight + 1e-7)
                area = 0
                weight = 0
                for i, box in enumerate(truck.boxes):
                    actual[box.item_number] += box.represented_qty
                    area += box.item_length * box.item_width
                    weight += box.item_weight * box.represented_qty
                    self.assertGreaterEqual(box.x, 0)
                    self.assertGreaterEqual(box.y, 0)
                    self.assertLessEqual(box.x + box.width, 1 + 1e-7)
                    self.assertLessEqual(box.y + box.height, 1 + 1e-7)
                    self.assertLessEqual(box.grouped_height + truck.roof_clearance_in, truck.truck_height)
                    for other in truck.boxes[i + 1:]:
                        overlap_x = min(box.x + box.width, other.x + other.width) - max(box.x, other.x)
                        overlap_y = min(box.y + box.height, other.y + other.height) - max(box.y, other.y)
                        self.assertFalse(overlap_x > 1e-7 and overlap_y > 1e-7)
                self.assertAlmostEqual(area, truck.total_used_pallets)
                self.assertAlmostEqual(weight, truck.total_weight)
                self.assertAlmostEqual(truck.utilization_percent, area / (truck.truck_length * truck.truck_width) * 100)
            self.assertEqual(actual, {r.item_number: r.qty for r in rows})
            self.assertAlmostEqual(sum(t.total_weight for t in trucks), sum(r.qty * r.item_weight for r in rows))
