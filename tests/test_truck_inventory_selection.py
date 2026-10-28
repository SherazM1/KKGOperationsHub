"""Streamlit workflow checks for parse refresh and item selection."""
import io
import unittest

import pandas as pd
from streamlit.testing.v1 import AppTest


class SelectionTests(unittest.TestCase):
    def test_bulk_double_stack_dimension_group(self):
        workbook = io.BytesIO()
        pd.DataFrame([
            {'STORE #': '1', 'DC #': '6006', 'ITEM #': item, 'QUANTITY': 1,
             'PALLET LENGTH': 48, 'PALLET WIDTH': 48, 'PALLET HEIGHT': 40, 'WEIGHT': 100}
            for item in ('A', 'B')
        ]).to_excel(workbook, index=False)
        workbook.name = 'group_stacking.xlsx'
        workbook.seek(0)
        app = AppTest.from_string('from app.ui.truck_inventory import render_truck_inventory_view\nrender_truck_inventory_view()')
        app.session_state['truck_combined_file'] = workbook
        app.run()
        app.button(key='truck_process_btn').click().run()
        group_selector = next(control for control in app.multiselect if control.label.startswith('Dimension groups'))
        group_selector.set_value([(48.0, 48.0, 40.0)])
        next(button for button in app.button if button.label == 'Apply to dimension groups').click().run()
        self.assertFalse(app.exception)
        draft = app.session_state['truck_item_setup_draft_by_load']['DC-6006']
        self.assertTrue(all(row['Is Stackable?'] == 'Yes' and row['Stack Qty'] == 2 for row in draft))
        self.assertFalse(app.session_state['truck_item_setup_confirmed_by_load']['DC-6006'])
        next(button for button in app.button if button.label == 'Save Item Setup').click().run()
        next(button for button in app.button if button.label == 'Build / Evaluate Truck Plan').click().run()
        self.assertFalse(app.exception)
        boxes = [box for truck in app.session_state['truck_trucks'] for box in truck.boxes]
        self.assertEqual(len(boxes), 1)
        self.assertEqual(sum(box.represented_qty for box in boxes), 2)
        self.assertEqual({member.item_number for member in boxes[0].allocations}, {'A', 'B'})
        app.radio(key='truck_workspace').set_value('Review & export').run()
        self.assertFalse(app.exception)
        self.assertTrue(any('Double stacks: 1 positions' in message.value for message in app.info))

    def test_guided_plan_review_and_settings_invalidation(self):
        workbook = io.BytesIO()
        pd.DataFrame([{'STORE #': '1', 'DC #': '6006', 'ITEM #': 'ITEM-A',
                       'QUANTITY': 2, 'PALLET LENGTH': 48, 'PALLET WIDTH': 48,
                       'PALLET HEIGHT': 30, 'WEIGHT': 100}]).to_excel(workbook, index=False)
        workbook.name = 'store.xlsx'
        workbook.seek(0)
        app = AppTest.from_string('from app.ui.truck_inventory import render_truck_inventory_view\nrender_truck_inventory_view()')
        app.session_state['truck_combined_file'] = workbook
        app.run()
        app.button(key='truck_process_btn').click().run()
        self.assertFalse(app.exception)
        next(button for button in app.button if button.label == 'Save Item Setup').click().run()
        next(button for button in app.button if button.label == 'Build / Evaluate Truck Plan').click().run()
        self.assertFalse(app.exception)
        self.assertTrue(app.session_state['truck_trucks'])
        next(button for button in app.button if button.label == 'Review truck plan').click().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.radio(key='truck_workspace').value, 'Review & export')
        app.radio(key='truck_workspace').set_value('Source data').run()
        self.assertFalse(app.exception)
        app.radio(key='truck_workspace').set_value('Plan').run()
        self.assertEqual(app.multiselect(key='truck_item_selector').value, ['ITEM-A'])
        app.selectbox(key='truck_preset').select('straight_box_truck_26_foot').run()
        self.assertFalse(app.exception)
        self.assertFalse(app.session_state['truck_trucks'])
        self.assertEqual(app.number_input(key='truck_operational_weight_threshold').value, 8700)

    def test_parse_once_then_select_dc_and_items(self):
        workbook = io.BytesIO()
        pd.DataFrame([
            {'STORE #': '1', 'DC #': '6006', 'ITEM #': item, 'QUANTITY': 1,
             'PALLET LENGTH': 48, 'PALLET WIDTH': 48, 'PALLET HEIGHT': 30, 'WEIGHT': 100}
            for item in ['ITEM-A', 'ITEM-B']
        ]).to_excel(workbook, index=False)
        workbook.name = 'store.xlsx'
        workbook.seek(0)
        app = AppTest.from_string('''
from app.ui.truck_inventory import _initialize_truck_state, _render_load_selection, render_inputs_tab
import streamlit as st
_initialize_truck_state()
if st.session_state.truck_normalized_records:
    _render_load_selection()
render_inputs_tab()
''')
        app.session_state['truck_combined_file'] = workbook
        app.run()
        app.button(key='truck_process_btn').click().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.multiselect(key='truck_loads_selector').value, ['DC-6006'])
        self.assertEqual(app.multiselect(key='truck_item_selector').value, ['ITEM-A', 'ITEM-B'])
        app.multiselect(key='truck_item_selector').set_value(['ITEM-A']).run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.session_state['truck_normalized_records']), 2)
        self.assertEqual([row['Item #'] for row in app.session_state['truck_item_setup']], ['ITEM-A'])
        app.multiselect(key='truck_item_selector').set_value([]).run()
        self.assertFalse(app.exception)
        self.assertEqual(app.multiselect(key='truck_item_selector').value, [])
