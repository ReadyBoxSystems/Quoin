# IMPORT_TEST_FILES.md

These `.xlsx` files are root-level fixtures for manually testing Quoin's Excel importer.

Use the running app at:

```text
http://localhost:3000
```

For each import, check:

- A new local configuration is created.
- Existing configurations are not overwritten.
- The **Sheets** strip appears between the grid and cell inspector.
- The expected Sheet names appear in the **Sheets** strip.
- Switching Sheets changes the visible grid without losing edits made on another Sheet.
- Grid dimensions expand enough for imported cells.
- Formula cells show the imported `=...` formula in the formula bar.
- Inspector shows Smart Cell names where expected.
- Review items are visible when expected.
- Runner Preview stays mostly empty unless imported named cells are later surfaced manually.

## Files

### 1. Basic Values

File:

```text
import-test-01-basic-values.xlsx
```

Expected:

- One Sheet: `Basic Values`
- Plain labels, numbers, boolean, and date values appear in the grid.
- No formulas.
- No review items.

### 2. Basic Formulas

File:

```text
import-test-02-basic-formulas.xlsx
```

Expected:

- One Sheet: `Basic Formulas`
- Formula cells import with visible `=...` formulas.
- `B5`, `B6`, and `B7` should calculate.
- No review items.

### 3. Named Cells

File:

```text
import-test-03-named-cells.xlsx
```

Expected:

- One Sheet: `Named Cells`
- Safe names promote to Smart Cells:
  - `design_span`
  - `design_plf`
- Named range `input_block` becomes a review item.
- Unsafe name `Design Span Unsafe` becomes a review item.
- Formula using names remains visible.

### 4. Multi-Sheet Cross-Sheet

File:

```text
import-test-04-multi-sheet-cross-sheet.xlsx
```

Expected:

- Two Sheets appear:
  - `Inputs`
  - `Calculator`
- Choose `Inputs` or `Calculator` in **Open Sheet** before import, then confirm the chosen Sheet opens first.
- Switching Sheets should reveal both imported grids.
- `input_span` promotes on `Inputs`.
- `total_load` promotes on `Calculator`.
- Cross-sheet formulas on `Calculator` remain visible and produce review warnings.
- Cross-sheet formulas are not expected to calculate across Sheets yet.

### 5. Reference Table Style

File:

```text
import-test-05-reference-table-style.xlsx
```

Expected:

- One Sheet: `Reference Table`
- Visible lookup-table-like grid imports.
- `selected_span` promotes to a Smart Cell.
- Named range `beam_table` shows as a review item.
- This remains the fixture for future reference-table direction.

### 6. Review Warnings

File:

```text
import-test-06-review-warnings.xlsx
```

Expected:

- One Sheet: `Review Warnings`
- Formulas import unchanged.
- Review list should flag:
  - cross-sheet reference
  - structured table reference
  - external workbook reference
  - semicolon argument separator
  - spill marker / dynamic array marker
- This file is meant to produce warnings.

### 7. Workbook Sheets

File:

```text
import-test-07-workbook-sheets.xlsx
```

Expected:

- Three Sheets appear:
  - `Inputs`
  - `Reference Data`
  - `Calculator`
- All Sheets should be available from the right-side **Sheets** strip after one import.
- `design_span` and `design_plf` promote on `Inputs`.
- `total_load` promotes on `Calculator`.
- Named range `beam_reference` appears as a review item.
- `Calculator` formulas with `Inputs!...` references stay visible and produce cross-sheet review warnings.

### 8. Open Selected Sheet

File:

```text
import-test-08-open-selected-sheet.xlsx
```

Manual step:

- In the import panel, choose `Calculator` as **Open Sheet** before clicking **Import Workbook**.

Expected:

- Three Sheets appear:
  - `Notes`
  - `Calculator`
  - `Lookup`
- `Calculator` opens first.
- Switching to `Notes` and `Lookup` shows their imported grids.
- `part_width`, `part_height`, and `part_area` promote on `Calculator`.
- `B5` on `Calculator` calculates from same-Sheet references.
- No review items expected.

### 9. Names Across Sheets

File:

```text
import-test-09-names-across-sheets.xlsx
```

Expected:

- Three Sheets appear:
  - `Geometry`
  - `Loads`
  - `Summary`
- Names promote on the correct Sheets:
  - `span_ft`
  - `member_depth`
  - `total_plf`
  - `summary_total_load`
- Same-Sheet formula on `Loads` should calculate.
- Cross-sheet formula on `Summary` should remain visible and produce review warnings.

### 10. Sheet Names With Spaces

File:

```text
import-test-10-sheet-names-with-spaces.xlsx
```

Expected:

- Two Sheets appear:
  - `Input Data`
  - `Calc Sheet`
- Sheet names with spaces display cleanly in the **Sheets** strip.
- `input_quantity` promotes on `Input Data`.
- `total_weight` promotes on `Calc Sheet`.
- Quoted cross-sheet formulas remain visible and produce review warnings.

### 11. Sparse Large Sheet

File:

```text
import-test-11-sparse-large-sheet.xlsx
```

Expected:

- One Sheet: `Sparse Large`
- Grid expands far enough to include columns through at least `M` and rows through at least `30`.
- Near cells around `A1:B3` import.
- Far cells around `K29:M30` import.
- `far_input` and `far_total` promote as Smart Cells.
- Formula `M30` remains visible and should calculate from same-Sheet references.

### 12. Workbook Review Mix

File:

```text
import-test-12-workbook-review-mix.xlsx
```

Expected:

- Two Sheets appear:
  - `Clean Inputs`
  - `Warning Formulas`
- `clean_input_a` promotes on `Clean Inputs`.
- `warning_cross_sheet_total` promotes on `Warning Formulas`.
- Named range `clean_input_block` appears as a review item.
- Review list aggregates warnings from the workbook:
  - cross-sheet reference
  - structured table reference
  - semicolon argument separator

## Regenerate

```bash
npm run fixtures:generate
```

## Smoke Test

```bash
npm run fixtures:smoke
```
