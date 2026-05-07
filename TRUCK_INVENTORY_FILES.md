# Truck Inventory Module - File Reference

## Models (Data Structures)
| File | Purpose | Key Classes |
|------|---------|------------|
| `app/models/truck_inventory_record.py` | Normalized inventory data model | `TruckInventoryRecord` |
| `app/models/truck_summary.py` | Truck assignment summary models | `TruckSummary`, `BoxLayout` |

## Services (Business Logic)
| File | Purpose | Key Functions |
|------|---------|---------------|
| `app/services/truck_inventory_parser.py` | Parse Excel files (PURE, CDW) | `parse_excel_file()`, `parse_combined_load_sheet()`, `detect_file_type()` |
| `app/services/truck_inventory_normalizer.py` | Convert raw rows to standard schema | `normalize_rows()`, `normalize_pure_row()`, `normalize_cdw_row()` |
| `app/services/truck_inventory_validator.py` | Validate records and flag issues | `validate_records()`, `get_validation_summary()` |
| `app/services/truck_inventory_pallet_calculator.py` | Calculate pallets (4 modes) | `calculate_pallets()`, `get_pallet_summary()`, and mode-specific functions |
| `app/services/truck_inventory_truck_assigner.py` | Assign loads to trucks sequentially | `assign_to_trucks()`, `get_truck_summary_stats()` |
| `app/services/truck_inventory_visualizer.py` | Render truck visualizations | `visualize_truck()`, `render_truck_visualization()`, `create_legend()` |
| `app/services/truck_inventory_export.py` | Export to CSV formats | `export_normalized_data_csv()`, `export_pallet_summary_csv()`, `export_truck_assignments_csv()`, `export_truck_boxes_csv()` |

## Configuration
| File | Purpose | Key Contents |
|------|---------|--------------|
| `app/utils/truck_presets.py` | Truck presets & color palette | `TRUCK_PRESETS`, `LOAD_GROUP_COLORS`, helper functions |

## UI
| File | Purpose | Key Functions |
|------|---------|---------------|
| `app/ui/truck_inventory.py` | Main Streamlit interface | `render_truck_inventory_view()`, 6 tab render functions |

## Integration
| File | Changes | Details |
|------|---------|---------|
| `app/main.py` | 3 changes | Import, enable button, add routing |

## Data Flow

```
User Uploads Files
       ↓
Parser (detect format + parse Excel)
       ↓
Normalizer (convert to standard schema)
       ↓
Validator (check required fields)
       ↓
Pallet Calculator (apply calculation mode)
       ↓
Pallet Summary (group & aggregate)
       ↓
Truck Assigner (sequential fill)
       ↓
Visualizer (render layouts) + Exporter (CSV output)
       ↓
User Reviews Tabs & Downloads
```

## Configuration & Customization Points

### 1. Pallet Calculation (HIGHEST PRIORITY for next build)
**File:** `app/services/truck_inventory_pallet_calculator.py`

Currently supports 4 modes - once business rules are finalized, update here:
```python
def calculate_pallets(records, mode="direct_qty", cases_per_pallet=40):
    # Replace or extend with actual business logic
```

### 2. Truck Assignment Algorithm
**File:** `app/services/truck_inventory_truck_assigner.py`

Currently uses simple sequential fill. To upgrade:
- Replace `assign_to_trucks()` logic
- Consider bin-packing algorithm or constraint solver
- Add weight/space distribution checks

### 3. Truck Presets
**File:** `app/utils/truck_presets.py`

Add custom presets:
```python
TRUCK_PRESETS = {
    "custom_key": TruckPreset(name="Custom Truck", pallet_capacity=20, ...),
}
```

### 4. Color Scheme
**File:** `app/utils/truck_presets.py`

Modify `LOAD_GROUP_COLORS` list to change visualization colors.

### 5. UI Layout
**File:** `app/ui/truck_inventory.py`

All Streamlit components are in this file - easy to modify tabs, columns, buttons, etc.

## Dependencies

### External Libraries (already in requirements.txt)
- `streamlit` - UI framework
- `pandas` - Data manipulation
- `openpyxl` - Excel reading
- `matplotlib` - Visualization

### Internal Dependencies
- Models depend on nothing (pure dataclasses)
- Services depend on models
- UI depends on services and models
- Main.py depends on UI

## Common Operations

### Add a new pallet calculation mode:
1. Create `estimate_pallets_my_mode()` function in `truck_inventory_pallet_calculator.py`
2. Add to `PALLET_CALC_MODES` dictionary
3. Add dispatch in `calculate_pallets()` function
4. UI automatically picks up the new mode from dropdown

### Add a new truck preset:
1. Create `TruckPreset` instance in `truck_presets.py` under `TRUCK_PRESETS`
2. UI automatically includes in truck preset dropdown

### Change visualization layout:
1. Modify `visualize_truck()` in `truck_inventory_visualizer.py`
2. Change box positioning logic (currently `x`, `y`, `width`, `height`)

### Add a new export format:
1. Create `export_my_format_csv()` function in `truck_inventory_export.py`
2. Add button in Export tab in `truck_inventory.py`

## Session State Variables

All prefixed with `truck_` to avoid conflicts:

```python
st.session_state.truck_pure_file          # Uploaded PURE file
st.session_state.truck_cdw_file           # Uploaded CDW file
st.session_state.truck_combined_file      # Uploaded combined file
st.session_state.truck_raw_records        # Raw parsed data
st.session_state.truck_normalized_records # After normalization & pallet calc
st.session_state.truck_validated_records  # After validation
st.session_state.truck_trucks             # Truck assignments
st.session_state.truck_pallet_summary     # Pallet summary by group
st.session_state.truck_grouping_rule      # Selected grouping rule
st.session_state.truck_pallet_mode        # Selected pallet calc mode
st.session_state.truck_preset             # Selected truck preset
st.session_state.truck_cases_per_pallet   # Cases per pallet config
```

---

**Quick Start for Modifications:**
1. **Change business rules?** → Update `truck_inventory_pallet_calculator.py`
2. **Change truck types?** → Update `truck_presets.py`
3. **Change UI?** → Update `truck_inventory.py`
4. **Change assignment logic?** → Update `truck_inventory_truck_assigner.py`
5. **Change visualization?** → Update `truck_inventory_visualizer.py`
