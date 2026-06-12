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
2. Choose the worksheet to turn into a Quoin grid.
3. Review imported values, formulas, names, and warning items.
4. Promote/adjust Smart Cells in Quoin.
5. Surface the runner-facing cells.

## Normal Cells

Normal cells:

- have coordinates like `B1`
- can hold plain values
- can hold formulas
- can reference other cells by coordinate
- can participate in Excel-style range formulas such as `SUM(A1:A5)`
- can be copied or filled down with coordinate references adjusted relative to the target cell
- are invisible to Runner Preview
- have no annotation, role, or surface behavior until named

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

## Spreadsheet UX

Current Excel-familiar behavior:

- `.xlsx` import for one worksheet at a time, creating a new local configuration
- imported cell values and formulas preserved in the grid
- workbook-defined safe single-cell names can promote cells to Smart Cells
- imported named ranges and formula compatibility issues appear as review items
- range formulas such as `SUM(A1:A5)` and rectangular ranges such as `SUM(A1:B3)`
- `IF(...)` formulas
- common uppercase function aliases such as `SUM`, `AVERAGE`, `MAX`, `MIN`, and `ROUND`
- undo/redo through toolbar buttons and `Ctrl+Z`, `Ctrl+Y`, `Ctrl+Shift+Z`
- internal copy/paste and fill-down that adjust coordinate references
- external spreadsheet paste into the normal grid
- Excel-style tabular paste into lookup tables
- add/delete row and add/delete column controls
- inspector dependency tracing for upstream `Depends On` and downstream `Used By`

When rows or columns are inserted/deleted, coordinate references should shift. References to deleted rows/columns become `#REF!`. Smart Cell names should remain stable.

## Excel Import

Current import behavior:

- user clicks `Import Excel`
- user selects an `.xlsx` file
- Quoin reads workbook sheets and names
- user chooses one worksheet
- Quoin imports values and formulas into a new local configuration
- formulas remain visible even if the engine cannot calculate them
- safe workbook-defined names pointing to a single cell can become Smart Cell names
- named ranges, unsafe names, cross-sheet references, external workbook references, structured table references, spill markers, and semicolon separators are reported as review items

Current importer files:

- `lib/import/types.ts` - neutral import model
- `lib/import/read.ts` - ExcelJS workbook reader
- `lib/import/convert.ts` - imported worksheet to Quoin grid converter
- `lib/import/convert.test.ts` - converter tests
- `lib/import/read.test.ts` - reader tests
- `scripts/generate-import-fixtures.mjs` - creates root `.xlsx` fixtures
- `scripts/smoke-import-fixtures.mjs` - reads fixtures through the importer
- `IMPORT_PLAN.md` - import scope and decisions
- `IMPORT_TEST_FILES.md` - manual fixture guide

Import intentionally does not clone all Excel behavior. It imports the calculation surface first; roles, surfacing, runner labels, lookup behavior, validation, compliance, and actions remain Quoin-native work after import.

## Role Semantics

Input:

- value supplied by admin/runner
- can be surfaced as a runner field
- can optionally expose dropdown options in Runner Preview

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

Runner Preview should only show surfaced Smart Cells and rule messages.

Normal coordinate cells should never leak into the runner-facing form.

Current Runner Preview grouping:

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
- Import Excel
- rename via the configuration name field
- switch via the configuration selector
- save sheet row/column dimensions

Unsaved changes remain in memory/local UI until the user clicks Save.

## Reference Table Direction

Lookup tables should move toward a workbook/reference-data model:

- The main grid remains the calculator/workflow authoring surface.
- Lookup data can live as a visible/imported reference table, similar to an Excel tab or named range.
- CSV ingestion is a minimum expected capability for workplace knowledge tables.
- Pasted grid ranges should eventually be promotable to named reference tables.
- Lookup Smart Cells should bind to a reference table, define match columns, and choose an output column.
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
- Local configuration storage keys currently live in the component.
- Local configurations include sheet dimensions for row/column changes.
- Excel import creates a new local configuration by default instead of overwriting the current one.
- Changing the storage shape may require a new local storage key, `Load Demo`, `Clear Sheet`, or creating a new local configuration.
