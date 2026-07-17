# DEV_NOTES.md

## Core Mental Model

Quoin should allow an admin to build a traditional spreadsheet first, then promote important cells into Smart Cells.

The intended workflow:

1. Build spreadsheet logic using normal cells.
2. Verify the math.
3. Name important cells.
4. Add display labels, annotations, roles, lookup/action/rule behavior.
5. Surface runner-facing cells.
6. Preview the generated runner form.
7. Save the work as a browser-local configuration.

Imported Excel calculators should follow the same mental model:

1. Import the workbook calculation surface.
2. Preserve workbook Sheets and choose the Sheet to open first.
3. Review imported values, formulas, names, and warning items.
4. Promote/adjust Smart Cells in Quoin.
5. Surface the runner-facing cells.

Supported imported Excel dropdown inputs are an intentional exception to fully manual Smart Cell promotion: Quoin promotes them to surfaced input Smart Cells during conversion so the dropdown behavior is visible in the grid, inspector, and Runner Preview immediately after import.

Exact spreadsheet lookup formulas are not an exception to the Smart Cell model. They remain normal formula behavior: `VLOOKUP` and `XLOOKUP` can calculate in coordinate cells without promoting the cell to a lookup Smart Cell.

## Normal Cells

Normal cells:

- have coordinates like `B1`
- can hold plain values
- can hold formulas
- can reference other cells by coordinate
- can participate in Excel-style range formulas such as `SUM(A1:A5)`
- can use exact-match `VLOOKUP`, exact/default `XLOOKUP`, cross-Sheet lookup ranges, and helper-key concatenation patterns
- can be copied or filled down with coordinate references adjusted relative to the target cell
- are invisible to Runner Preview
- have no annotation, role, or surface behavior until named
- can still render as dropdowns in the grid if they carry imported/configured `inputOptions`, but Runner Preview requires surfaced Smart Cells

## Smart Cells

A cell becomes smart when it gets a name.

Smart Cells can have:

- name
- display label
- role
- type
- annotation
- runner-facing rule message
- surface toggle
- input dropdown options
- lookup config
- validation/compliance behavior

Smart Cells can be referenced by name in formulas.
Formula copy/fill should not duplicate Smart Cell names; named references should remain semantic and unchanged.

Smart Cell names should remain formula/backend-safe, for example `design_span` or `recommended_beam`.
Display labels are for runner-facing text, for example `Design Span (ft)` or `Recommended Beam`.

In multi-Sheet configurations, Smart Cell names are workbook-scoped. A formula on any Sheet can reference a unique Smart Cell name directly, such as `=design_span * design_plf`, without using coordinate-style cross-Sheet syntax.

If a Smart Cell has dropdown options, the grid should render it as a dropdown cell and Runner Preview should render the same controlled option set. The inspector remains the configuration surface for editing short embedded option lists. Imported Excel dropdowns with supported same-workbook range or named-range sources are snapshotted into embedded options for now. Longer dropdown option sets should eventually use live visible reference data, named ranges/tables, or CSV-ingested datasets instead of large cell-local text lists.

## Spreadsheet UX

Current Excel-familiar behavior:

- `.xlsx` import creates one local configuration with preserved Sheets
- Sheet switching uses a tab strip directly below the formula bar and above the grid
- default grid dimensions and cell sizing are compact enough to show a broader working area
- mouse cell selection and double-click editing are stable enough for normal spreadsheet clicking
- imported cell values and formulas preserved in the grid
- workbook-defined safe single-cell names can promote cells to Smart Cells
- imported merged ranges, named ranges, and formula compatibility issues appear as review items
- basic imported cross-Sheet formulas calculate across preserved Sheets
- unique Smart Cell names can be referenced directly from formulas on any Sheet in the workbook
- exact lookup formulas calculate as normal formulas, including cross-Sheet ranges and concatenated helper keys
- Runner Preview gathers surfaced Smart Cells across the full workbook and groups them by Sheet
- Runner Preview hides optional sections when there are no surfaced cells for that section
- range formulas such as `SUM(A1:A5)` and rectangular ranges such as `SUM(A1:B3)`
- `IF(...)` formulas
- common uppercase function aliases such as `SUM`, `AVERAGE`, `MAX`, `MIN`, and `ROUND`
- `ROUNDUP`, exact-match `VLOOKUP`, and exact/default `XLOOKUP`
- undo/redo through toolbar buttons and `Ctrl+Z`, `Ctrl+Y`, `Ctrl+Shift+Z`
- internal copy/paste and fill-down that adjust coordinate references
- external spreadsheet paste into the normal grid
- Excel-style tabular paste into lookup tables
- Smart Cell dropdown options render as dropdown controls in the grid and Runner Preview
- simple imported Excel data-validation lists and supported same-workbook range sources can become Smart Cell dropdown options
- supported imported Excel dropdown cells are promoted to surfaced input Smart Cells, using a safe generated name and nearby label text when no workbook-defined name exists
- add/delete row and add/delete column controls
- Sheet rename support updates direct cross-Sheet references that use the renamed Sheet
- sticky row and column headers support large grid navigation
- inspector dependency tracing for upstream `Depends On` and downstream `Used By`

When rows or columns are inserted/deleted, coordinate references should shift. References to deleted rows/columns become `#REF!`. Smart Cell names should remain stable.

## Excel Import

Current import behavior:

- user clicks `Import Workbook`
- user selects an `.xlsx` file
- Quoin reads workbook sheets and names
- user chooses which Sheet to open first
- Quoin imports workbook Sheets into a new local configuration
- formulas remain visible even if the engine cannot calculate them
- safe workbook-defined names pointing to a single cell can become Smart Cell names
- supported Excel data-validation dropdown cells become surfaced input Smart Cells
- exact lookup formulas are preserved as normal formula cells and can calculate without lookup Smart Cell conversion
- approximate lookup behavior remains review-first
- merged Excel ranges are reported as review items; only the top-left cell is imported and covered cells are left blank
- named ranges, unsafe names, external workbook references, structured table references, spill markers, and semicolon separators are reported as review items
- simple typed Excel data-validation lists are imported as embedded dropdown options
- supported same-workbook Excel data-validation range sources are snapshotted as embedded dropdown options
- supported workbook-defined named range dropdown sources are snapshotted as embedded dropdown options
- unsupported formula/table/external Excel data-validation list sources are review items until reference-data-backed dropdowns exist

Current importer files:

- `lib/import/types.ts` - neutral import model
- `lib/import/read.ts` - ExcelJS workbook reader
- `lib/import/convert.ts` - imported worksheet to Quoin grid converter
- `lib/import/convert.test.ts` - converter tests
- `lib/import/read.test.ts` - reader tests
- `scripts/generate-import-fixtures.mjs` - creates root `.xlsx` fixtures
- `scripts/smoke-import-fixtures.mjs` - reads fixtures through the importer
- `docs/archive/import-testing/` - historical import plan, manual fixture guide, manual test notes, and generated workbook fixtures from the completed importer pass

Import intentionally does not clone all Excel behavior. It imports the calculation surface first; roles, surfacing, runner labels, lookup behavior, validation, compliance, and actions remain Quoin-native work after import.

Current dropdown import translation:

- literal Excel list validations are parsed into options
- same-Sheet, cross-Sheet, quoted-Sheet, and workbook-defined named range list validations are snapshotted into options
- blank source values are skipped and duplicate option values are deduplicated in first-seen order
- imported dropdown cells are converted to surfaced input Smart Cells
- unsupported table/formula/external list sources produce review items
- exact `VLOOKUP` and `XLOOKUP` formulas are preserved and calculated as formulas; they are not converted into lookup Smart Cells

## Role Semantics

Input:

- value supplied by admin/runner
- can be surfaced as a runner field
- can optionally expose dropdown options in both the grid and Runner Preview

Formula:

- calculated value
- can be internal or surfaced

Output:

- calculated or entered result
- can be surfaced as a runner result

Lookup:

- calculated from a lookup table
- can match on multiple input criteria
- lookup misses should show `#ERR`, not blank the sheet
- current lookup tables are embedded in Smart Cell metadata as a prototype shortcut
- intended direction is for lookup Smart Cells to query first-class reference tables / CSV-ingested datasets
- separate from normal spreadsheet lookup formulas; use lookup Smart Cells when the lookup needs named criteria, runner surfacing, editable table behavior, or clearer audit structure

Action:

- runner-facing shop note or required action
- can be powered by lookup behavior
- should appear in Runner Preview as a shop action, not just a generic result

Validation:

- displays `PASS`, `FAIL`, or `#ERR`
- false condition means the run failed validation
- math should still run where possible

Compliance:

- displays `OK`, `WARN`, or `#ERR`
- true condition means warn the runner
- warning does not invalidate the run

## Runner Preview

Runner Preview should only show surfaced Smart Cells and rule messages across the current workbook.

Normal coordinate cells should never leak into the runner-facing form.

Current Runner Preview grouping, with Sheet headings when multiple Sheets contribute surfaced cells:

- Inputs
- Outputs
- Shop Actions
- Review Flags
- Validation

## Local Configurations

Configurations are browser-local for now.

Available actions:

- New
- Save
- Duplicate
- Delete
- Import Workbook
- rename via the configuration name field
- switch via the configuration selector
- save sheet row/column dimensions

Unsaved changes remain in memory/local UI until the user clicks Save.

## Reference Table Direction

Lookup tables should move toward a workbook/reference-data model:

- The main grid remains the calculator/workflow authoring surface.
- Lookup data can live as a visible/imported reference table, similar to an Excel tab or named range.
- Dropdown option sources for longer lists should also be able to bind to visible/imported reference data.
- CSV ingestion is a minimum expected capability for workplace knowledge tables.
- Pasted grid ranges should eventually be promotable to named reference tables.
- Lookup Smart Cells should bind to a reference table, define match columns, and choose an output column.
- Dropdown Smart Cells should bind to a reference table/range option column for long option sets.
- Embedded lookup tables in the inspector are acceptable for the current prototype but should not be treated as the final UX.
- Later publish/run/audit work should preserve which reference table/version/row was used.

## Demo Configuration

The default demo is `Demo - Beam Selection`.

Rules:

- Demo data is fake/non-proprietary.
- Do not copy or reference proprietary shop spreadsheets or screenshots.
- The demo should show Quoin's workflow direction: enter drawing conditions, calculate context, return recommendations, show actions, and flag review conditions.

## Current Implementation Notes

- Main component: `components/variable-sheet.tsx`
- Engine: `lib/engine/index.ts`
- Engine tests: `lib/engine/engine.test.ts`
- Workbook/Sheet types: `lib/sheet/types.ts`
- Git remote: `https://github.com/ReadyBoxSystems/Quoin.git`
- Primary branch: `main`
- Baseline commit: `5259f17 Initial Quoin prototype baseline`
- Local configuration storage keys currently live in the component.
- Local configurations include one or more Sheets, an active Sheet id, and legacy top-level cell/dimension fields for compatibility.
- Excel import creates a new multi-Sheet local configuration by default instead of overwriting the current one.
- Workbook-scoped Runner Preview uses all current Sheets, including unsaved active-Sheet edits, while the authoring grid still shows one active Sheet at a time.
- Dropdown rendering in the grid checks `inputOptions.length > 0`; Runner Preview still depends on surfaced Smart Cells.
- Imported dropdown Smart Cell identity generation happens in `lib/import/convert.ts`.
- Lookup formula evaluation happens in `lib/engine/index.ts`; import formula review for lookup behavior happens in `lib/import/convert.ts`.
- Changing the storage shape may require a new local storage key, `Load Demo`, `Clear Sheet`, or creating a new local configuration.
