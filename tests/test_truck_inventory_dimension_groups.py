import unittest

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_item_setup import get_dimension_groups, apply_dimension_group_stacking


class DimensionGroupTests(unittest.TestCase):
    def rows(self):
        return [dict(zip(['Item #', 'Length', 'Width', 'Height', 'Weight'], values),
                     **{'Preset': 'Custom', 'Is Stackable?': 'No', 'Stack Qty': 1})
                for values in [('A', 48, 48, 40, 100), ('B', 48, 48, 40, 200),
                               ('C', 48, 48, 41, 100), ('HALF', 48, 48, 40, 100)]]

    def records(self):
        return [TruckInventoryRecord(source_type='test', source_file='', po_number='PO',
                                      item_number='HALF', pallet_type='Half')]

    def test_groups_use_all_three_dimensions_not_item_or_weight(self):
        groups = get_dimension_groups(self.rows())
        self.assertEqual(groups[(48, 48, 40)], ['A', 'B', 'HALF'])
        self.assertEqual(groups[(48, 48, 41)], ['C'])
        self.assertFalse(get_dimension_groups([{'Item #': 'X', 'Length': float('inf'), 'Width': 48, 'Height': 40}]))

    def test_only_selected_group_changes_including_half_pallets(self):
        original = self.rows()
        updated, skipped = apply_dimension_group_stacking(original, self.records(), [(48, 48, 40)], True)
        self.assertFalse(skipped)
        self.assertEqual([row['Is Stackable?'] for row in updated], ['Yes', 'Yes', 'No', 'Yes'])
        self.assertEqual([row['Stack Qty'] for row in updated], [2, 2, 1, 2])
        self.assertTrue(all(row['Is Stackable?'] == 'No' for row in original))
        for before, after in zip(original, updated):
            for key in ('Length', 'Width', 'Height', 'Weight'):
                self.assertEqual(before[key], after[key])

    def test_disable_group_and_half_pallet_preset(self):
        rows = self.rows()
        rows[0]['Preset'] = 'Half Pallet'
        updated, skipped = apply_dimension_group_stacking(rows, [], [(48, 48, 40)], True)
        self.assertFalse(skipped)
        self.assertEqual(updated[0]['Stack Qty'], 2)
        updated, skipped = apply_dimension_group_stacking(updated, [], [(48, 48, 40)], False)
        self.assertFalse(skipped)
        self.assertTrue(all(row['Stack Qty'] == 1 for row in updated))
