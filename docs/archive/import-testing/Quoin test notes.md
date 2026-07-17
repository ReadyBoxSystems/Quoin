Quoin import test notes

Use this sequence in the running app at http://localhost:3000.

General checks after each import:

- A new local configuration was created.
- Existing configurations were not overwritten.
- The Sheets strip appears between the grid and cell panel.
- Expected Sheet names appear in the Sheets strip.
- Switching Sheets changes the visible grid.
- Edits made on one Sheet are still there after switching away and back.
- Grid dimensions expanded enough for the imported cells.
- Formula bar shows the imported formula when selecting formula cells.
- Inspector shows Smart Cell names where expected.
- Runner Preview stays mostly empty unless imported named cells are later surfaced manually.

1. Basic Values

Import:

import-test-01-basic-values.xlsx

Expected:

- One Sheet: Basic Values
- Plain labels, numbers, boolean, and date values appear in the grid
- No formulas
- No review items

Result: PASS. NO ERRORS TO REPORT.


2. Basic Formulas

Import:

import-test-02-basic-formulas.xlsx

Expected:

- One Sheet: Basic Formulas
- Formula cells import with visible =... formulas
- B5, B6, B7 should calculate
- No review items expected

Result: PASS. NO ERRORS TO REPORT


3. Named Cells

Import:

import-test-03-named-cells.xlsx

Expected:

- Safe names promote to Smart Cells:
  - design_span
  - design_plf
- Named range input_block becomes a review item
- Unsafe name Design Span Unsafe becomes a review item
- Formula using names should remain visible

Result: PASS. NO ERRORS TO REPORT


4. Multi-Sheet Cross-Sheet

Import:

import-test-04-multi-sheet-cross-sheet.xlsx

Expected:

- Two Sheets: Inputs, Calculator
- Choose either Sheet in Open Sheet, then confirm that Sheet opens first
- Sheet strip lets you switch between Inputs and Calculator
- input_span promotes on Inputs
- total_load promotes on Calculator
- Cross-sheet formulas remain visible on Calculator
- Cross-sheet formulas produce review warnings
- Cross-sheet formulas are not expected to calculate across Sheets yet

Result: MOSTLY PASS. CONCERNING THIS SECTION: "Cross-sheet formulas are not expected to calculate across Sheets yet". WHILE TRUE, THE RUNNER ITSELF SHOULD SHOULD BOTH. CURRENTLY, THE RUNNER ONLY SHOWS THAT PARTICULAR SHEET'S SMART CELLS VS THE EXPECTED COMBINED CELLS. WHILE THE MATH/CONNECTIONS BETWEEN THE CELLS MIGHT NOT EXIST YET, THEY SHOULD SHOW ON THE RUNNER.


5. Reference Table Style

Import:

import-test-05-reference-table-style.xlsx

Expected:

- One Sheet: Reference Table
- Visible lookup-table-like grid imports
- selected_span promotes to a Smart Cell
- beam_table named range shows as a review item
- This is still the fixture for future reference-table direction

Result: PASS. NO ERRORS TO REPORT


6. Review Warnings

Import:

import-test-06-review-warnings.xlsx

Expected:

- One Sheet: Review Warnings
- Imports formulas unchanged
- Review list should flag:
  - cross-sheet reference
  - structured table reference
  - external workbook reference
  - semicolon argument separator
  - spill marker / dynamic array marker

Result: PASS. NO ERRORS TO REPORT


7. Workbook Sheets

Import:

import-test-07-workbook-sheets.xlsx

Expected:

- Three Sheets: Inputs, Reference Data, Calculator
- All three Sheets appear after one import
- design_span and design_plf promote on Inputs
- total_load promotes on Calculator
- beam_reference named range appears as a review item
- Calculator formulas with Inputs! references stay visible and produce cross-sheet review warnings

Result: FAILURE. SAME ISSUE AS TEST FOUR. THE RUNNER DOES NOT AGGREGATE THE SHEETS TO ONE FORM. EACH ONE IS SEPARATE, SO NOTHING ACTUALLY WORKS/CONNECTS ON THE RUNNER SIDE.


8. Open Selected Sheet

Import:

import-test-08-open-selected-sheet.xlsx

Before importing:

- Choose Calculator as Open Sheet

Expected:

- Three Sheets: Notes, Calculator, Lookup
- Calculator opens first
- Notes and Lookup are available in the Sheet strip
- part_width, part_height, and part_area promote on Calculator
- B5 on Calculator calculates from same-Sheet references
- No review items expected

Result: CHOOSING THE CALCULATOR AS THE SHEET TO OPEN PASSES. SAME RUNNER ISSUE AS FOUR AND SEVEN.


9. Names Across Sheets

Import:

import-test-09-names-across-sheets.xlsx

Expected:

- Three Sheets: Geometry, Loads, Summary
- Names promote on the correct Sheets:
  - span_ft
  - member_depth
  - total_plf
  - summary_total_load
- Same-Sheet formula on Loads should calculate
- Cross-sheet formula on Summary should remain visible and produce review warnings

Result: I SEE THE REVIEW WARNING. SAME RUNNER ISSUE AS FOUR, SEVEN, AND EIGHT.


10. Sheet Names With Spaces

Import:

import-test-10-sheet-names-with-spaces.xlsx

Expected:

- Two Sheets: Input Data, Calc Sheet
- Sheet names with spaces display cleanly in the Sheet strip
- input_quantity promotes on Input Data
- total_weight promotes on Calc Sheet
- Quoted cross-sheet formulas remain visible and produce review warnings

Result: PASS. SHEET NAMES WITH SPACES IMPORT FIME.


11. Sparse Large Sheet

Import:

import-test-11-sparse-large-sheet.xlsx

Expected:

- One Sheet: Sparse Large
- Grid expands through at least column M and row 30
- Near cells around A1:B3 import
- Far cells around K29:M30 import
- far_input and far_total promote as Smart Cells
- Formula M30 remains visible and should calculate from same-Sheet references

Result: PASS. NO ERRORS TO REPORT


12. Workbook Review Mix

Import:

import-test-12-workbook-review-mix.xlsx

Expected:

- Two Sheets: Clean Inputs, Warning Formulas
- clean_input_a promotes on Clean Inputs
- warning_cross_sheet_total promotes on Warning Formulas
- clean_input_block named range appears as a review item
- Review list aggregates warnings from the workbook:
  - cross-sheet reference
  - structured table reference
  - semicolon argument separator

Result: PASS OTHER THAN THE RUNNER ISSUE. SAME AS FOUR, SEVEN, EIGHT, AND NINE.


Summary / notes:


Workbook Runner / Cross-Sheet Calculation Retest

Use this sequence after the workbook-scoped Runner Preview and basic cross-Sheet calculation update.

General checks after each import:

- A new local configuration was created.
- Existing configurations were not overwritten.
- The Sheets strip appears between the grid and cell panel.
- Expected Sheet names appear in the Sheets strip.
- Switching Sheets changes the visible grid.
- Formula bar shows imported cross-Sheet formulas when selecting formula cells.
- Basic cross-Sheet formulas calculate across preserved Sheets.
- Cross-Sheet formulas should no longer appear as review warnings by themselves.
- Runner Preview gathers surfaced Smart Cells from all Sheets in the workbook.
- Runner Preview groups surfaced Smart Cells by Sheet when more than one Sheet contributes surfaced cells.
- Normal coordinate cells do not appear in Runner Preview.
- Unsurfaced Smart Cells do not appear in Runner Preview.
- Editing a Runner input from one Sheet updates that Sheet, even when another Sheet is active.


13. Workbook Runner Cross-Sheet

Import:

import-test-13-workbook-runner-cross-sheet.xlsx

Expected:

- Two Sheets: Runner Inputs, Runner Calculator
- runner_span and runner_plf promote on Runner Inputs
- runner_total_load and runner_review_note promote on Runner Calculator
- Runner Calculator formulas remain visible
- runner_total_load calculates from Runner Inputs values
- No cross-Sheet review warning expected
- Manually surface runner_span, runner_plf, runner_total_load, and runner_review_note
- Runner Preview shows surfaced cells from both Sheets, grouped by Sheet
- Editing runner_span or runner_plf in Runner Preview updates Runner Inputs
- runner_total_load updates in Runner Preview after input edits

Result:


14. Cross-Sheet Range

Import:

import-test-14-cross-sheet-range.xlsx

Expected:

- Two Sheets: Loads, Summary
- load_a promotes on Loads
- range_total and range_average promote on Summary
- SUM(Loads!B3:B5) calculates on Summary
- AVERAGE(Loads!B3:B5) calculates on Summary
- No cross-Sheet range review warning expected
- Manually surface load_a, range_total, and range_average
- Runner Preview shows surfaced cells from both Sheets, grouped by Sheet

Result:


15. Workbook Review After Cross-Sheet Support

Import:

import-test-15-workbook-review-after-cross-sheet-support.xlsx

Expected:

- Two Sheets: Supported Inputs, Review Items
- supported_input promotes on Supported Inputs
- supported_cross_sheet_total promotes on Review Items
- Supported cross-Sheet formula calculates
- Review list should not flag cross-Sheet references
- Review list should still flag:
  - structured table reference
  - external workbook reference
  - semicolon argument separator
- Manually surface supported_input and supported_cross_sheet_total
- Runner Preview shows surfaced cells from both Sheets, grouped by Sheet

Result:
