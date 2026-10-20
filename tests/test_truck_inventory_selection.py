"""Streamlit workflow checks for parse refresh and item selection."""
import io
import unittest

import pandas as pd
from streamlit.testing.v1 import AppTest


class SelectionTests(unittest.TestCase):
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
