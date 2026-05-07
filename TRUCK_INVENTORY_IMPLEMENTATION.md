# Truck Inventory Module - Implementation Summary

## Overview
The Truck Inventory module is a new operational planning tool integrated into the Kendal King Operations Hub. It enables users to upload PURE and CDW order files, normalize the data, calculate pallet counts, assign loads to trucks, visualize truck layouts, and export results for further analysis.

## What Was Added

### 1. Core Data Models (`app/models/`)
- **`truck_inventory_record.py`** - Normalized record model matching the standardized schema
  - Contains all 29 fields (source_type, po_number, item details, pallet calculations, validation status, etc.)
  - Includes conversion to dictionary for DataFrame operations

- **`truck_summary.py`** - Truck assignment and summary models
  - `BoxLayout` - Represents a colored load/pallet box within a truck
  - `TruckSummary` - Represents a truck with its capacity, utilization, and box contents

### 2. Configuration (`app/utils/`)
- **`truck_presets.py`** - Truck capacity presets and color schemes
  - 4 standard truck presets (53ft dry van, 48ft trailer, half truck, quarter truck)
  - Color palette for load groups (10 colors, cycling for unlimited groups)
  - Helper functions for color mapping and preset selection

### 3. Data Processing Services (`app/services/`)

#### Input & Parsing
- **`truck_inventory_parser.py`** - Parse Excel files (PURE, CDW, combined)
  - Detects file type based on filename and column content
  - Handles Excel reading with error management
  - Returns structured parse results

#### Data Normalization
- **`truck_inventory_normalizer.py`** - Normalize rows to standard schema
  - `normalize_pure_row()` - Convert PURE format to standard
  - `normalize_cdw_row()` - Convert CDW format to standard
  - `normalize_combined_row()` - Convert combined load sheets
  - Flexible column name detection for real-world file variations
  - Fallback detection when file type is unclear

#### Validation
- **`truck_inventory_validator.py`** - Validate normalized records
  - Check required fields (PO Number, Description/Item, Quantity)
  - Flag optional fields as warnings
  - Produce validation summary statistics
  - Per-record validation status and notes

#### Pallet Calculation
- **`truck_inventory_pallet_calculator.py`** - Calculate pallet counts (MVP with multiple modes)
  - **direct_qty**: Use quantity directly
  - **qty_divided_by_cases**: Divide by configurable cases-per-pallet (default 40)
  - **description_inference**: Heuristic inference from description keywords
  - **manual_override**: Use pre-set estimated_pallets field
  - Easy to replace/extend when business rules are finalized
  - Generates pallet summary by load group

#### Truck Assignment
- **`truck_inventory_truck_assigner.py`** - Sequential truck filling logic
  - Groups records by load_group, po_number, or none (configurable)
  - Fills trucks sequentially until capacity is reached
  - Creates new trucks on overflow
  - Calculates utilization percentages
  - Tracks remaining capacity

#### Visualization
- **`truck_inventory_visualizer.py`** - 2D top-down truck visualization
  - Matplotlib-based truck layout rendering
  - Color-coded load boxes showing pallet counts
  - Utilization overlay
  - Legend/key for all load groups
  - Streamlit integration for easy embedding

#### Export
- **`truck_inventory_export.py`** - Multi-format CSV exports
  - Normalized data export
  - Pallet summary export
  - Truck assignment summary export
  - Box placement details export
  - Combined report capability

### 4. UI Module (`app/ui/`)
- **`truck_inventory.py`** - Main Streamlit interface with 6 tabs
  
  **Tab 1: Inputs**
  - File uploaders for PURE, CDW, and combined load sheet files
  - Configuration controls:
    - Grouping rule (load_group, po_number, none)
    - Pallet calculation mode (4 options)
    - Truck preset selector
    - Cases-per-pallet configurable input
  - "Process Files" button to run full pipeline
  
  **Tab 2: Normalized Data**
  - Statistics overview (total records, source types, load groups, total pallets)
  - Full data table preview (scrollable, use_container_width)
  - Download button for normalized CSV
  
  **Tab 3: Pallet Summary**
  - Pallet summary table by load group (items, total pallets, total qty)
  - Summary statistics
  - Download button for pallet summary CSV
  
  **Tab 4: Truck Builder**
  - Overall truck assignment statistics (total trucks, capacity, utilization)
  - Truck-by-truck assignments with expandable details
  - Utilization metrics per truck
  - Load group breakdown per truck
  - Download buttons for truck and box CSV exports
  
  **Tab 5: Visualization**
  - 2D top-down truck visualization (matplotlib)
  - Color-coded boxes for each load group
  - Utilization and capacity overlay
  - Truck-by-truck layout
  - Legend showing all load groups and colors
  
  **Tab 6: Export**
  - All available export options consolidated
  - Download buttons for 4 CSV formats
  - Clear disabled state for exports with no data

### 5. Main Entry Point Updates (`app/main.py`)
- Imported `render_truck_inventory_view` from new module
- Enabled the "Truck Inventory" button (removed "Coming Soon", removed disabled=True)
- Added routing case in `main()` function to display the module
- Applied theme styling and hub header to match other tools

## How It Works (End-to-End)

1. **User Uploads Files** → Files are stored in session state
2. **Process Files Button** → Triggers pipeline:
   - Parse Excel files (auto-detect format)
   - Normalize rows to standard schema
   - Validate records (check required fields)
   - Calculate pallets using configured mode
   - Get pallet summary by group
   - Assign to trucks sequentially
   - All results stored in session state
3. **User Reviews Tabs** → Switch between inputs, normalized data, summaries, truck assignments, visualization, and export
4. **User Exports** → Download any of 4 CSV formats for further analysis

## Assumptions & Design Decisions

### 1. **File Type Detection**
- Files are detected by filename hints ("pure", "cdw" in filename) and column content analysis
- Falls back to heuristic pattern matching if unclear
- Allows user to re-process with different configurations without re-uploading

### 2. **Column Name Flexibility**
- Normalizer checks multiple possible column name variants
- Handles common variations (e.g., "PO", "PO_Number", "PO Number", "po_number")
- Real-world Excel files have inconsistent naming; this design handles that gracefully

### 3. **Pallet Calculation is Intentionally Modular**
- **Why?** Business rules for converting quantity to pallets are not yet finalized
- **Solution:** MVP supports 4 placeholder modes; easy to swap implementation later
- **Location:** All logic in `truck_inventory_pallet_calculator.py` with clear mode dispatch
- **Extension path:** Replace `calculate_pallets()` function once rules are known

### 4. **Sequential Truck Assignment**
- Groups are filled sequentially into trucks using simple greedy algorithm
- No bin-packing optimization (kept MVP simple)
- Fills trucks in order; creates new truck when current capacity exceeded
- Easy to upgrade to smarter assignment algorithms later

### 5. **Truck Visualization**
- 2D top-down layout (not 3D, not drag-and-drop)
- Boxes stack vertically in truck
- Color-coded by load group
- Good enough for MVP; extensible for future 3D or interactive layouts

### 6. **Validation Strategy**
- Marks records as "valid", "warning", or "error"
- Requires: PO Number, Description OR Item Number, Quantity
- Warnings: Missing optional fields (Delivery Date, Event Code for PURE)
- Users can proceed with warnings; errors indicate missing critical data

### 7. **Session State Management**
- All data stored in `st.session_state` (truck_* prefixed)
- Users can modify configuration and re-process without re-uploading
- Streamlit rerun on button click auto-updates all downstream tabs

## File Structure

```
app/
├── models/
│   ├── truck_inventory_record.py      # Normalized data model
│   └── truck_summary.py               # Truck & box models
├── services/
│   ├── truck_inventory_parser.py      # File parsing (PURE/CDW)
│   ├── truck_inventory_normalizer.py  # Data normalization
│   ├── truck_inventory_validator.py   # Validation logic
│   ├── truck_inventory_pallet_calculator.py  # Pallet calculation (4 modes)
│   ├── truck_inventory_truck_assigner.py     # Sequential assignment
│   ├── truck_inventory_visualizer.py  # Matplotlib visualization
│   └── truck_inventory_export.py      # CSV export functions
├── ui/
│   └── truck_inventory.py             # Main Streamlit UI (6 tabs)
├── utils/
│   └── truck_presets.py               # Truck presets & colors
└── main.py                            # Updated with routing
```

## Known Limitations & TODO for Next Phase

### Intentional Placeholders (MVP Design)
1. **Pallet Calculation Rules** - Currently 4 configurable modes; business rules unknown
   - Once finalized, update `truck_inventory_pallet_calculator.py`
   - Consider adding customer-specific rules or time-based overrides
   
2. **Truck Assignment Algorithm** - Currently simple sequential fill
   - No optimization for capacity utilization
   - Could upgrade to bin-packing or constraint-based solver
   - Consider splitting large shipments across multiple trucks
   
3. **Visualization** - 2D top-down layout only
   - No 3D rendering
   - No drag-and-drop rebalancing
   - No detailed spatial constraints (height, weight distribution)
   - Good enough for MVP; can extend with matplotlib 3D or interactive library

4. **Load Group Inference** - Currently uses Event Code (PURE) or manual entry
   - Could add ML-based categorization
   - Could tie to customer master data
   - Could infer from item description or category

### Minor Enhancements for Next Phase
1. Add date range filters (delivery date, PO date)
2. Add search/filter capability on normalized data table
3. Add custom truck preset creation in UI
4. Add drag-and-drop file uploaders (Streamlit 1.35+)
5. Add bulk editing of load groups
6. Add truck capacity override per upload
7. Add delivery address geocoding for distance optimization
8. Add multi-file batch processing with progress tracking
9. Add comparison mode (before/after configurations)
10. Add PDF report generation with truck layouts

## Testing Checklist

- [x] All Python syntax validates (no errors)
- [x] Imports resolve correctly
- [x] Models instantiate properly
- [x] Services have no circular dependencies
- [ ] Streamlit app starts without errors (manual test)
- [ ] File upload works (manual test)
- [ ] File parsing handles empty/malformed Excel (manual test)
- [ ] Normalization produces correct schema (manual test)
- [ ] Validation assigns correct status (manual test)
- [ ] Pallet calculation works in all 4 modes (manual test)
- [ ] Truck assignment fills sequentially (manual test)
- [ ] Visualization renders without errors (manual test)
- [ ] CSV exports are valid (manual test)
- [ ] Tabs switch properly (manual test)
- [ ] Back button returns to home (manual test)

## Next Build Steps (Recommended Priority)

1. **Manual Testing** - Run Streamlit app with sample PURE/CDW files
2. **Sample Data Creation** - Create realistic PURE and CDW test files
3. **Business Rule Finalization** - Confirm pallet calculation rules with stakeholders
4. **Rule Implementation** - Once rules are known, update `truck_inventory_pallet_calculator.py`
5. **Advanced Assignment** - If needed, upgrade truck assignment algorithm
6. **Customer Master Data** - Link load groups to customer master data
7. **Reporting** - Add PDF generation with truck layouts
8. **Performance** - Optimize for large file uploads (1000+ rows)

## Integration Notes

- Does NOT break existing tools (Label Maker, BOL Generator, SKID Tags)
- Uses same theming system as other Operations Hub tools
- Follows same import and service patterns
- Session state isolated with "truck_" prefix to avoid conflicts
- Back button navigates consistently with other modules

## Questions for Stakeholders

1. What is the final rule for quantity → pallet count conversion?
2. Should certain item categories always occupy a full truck?
3. Are there weight/size constraints beyond pallet count?
4. Should load groups be automatically assigned or always manual?
5. Do certain customers require specific truck types?
6. Is there a preferred delivery date clustering strategy?
7. Should the system support partial pallets or always round up?

---

**Created:** May 7, 2026  
**Status:** MVP Ready for Testing  
**Branch:** main
