"""Streamlit workflow checks for parse refresh and item selection."""
import io
import unittest

import pandas as pd
from streamlit.testing.v1 import AppTest


class SelectionTests(unittest.TestCase):
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
