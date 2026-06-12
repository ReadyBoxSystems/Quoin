# ROADMAP.md

## Phase 1 - Spreadsheet UX

Goal: make the admin authoring surface feel like Excel with a smart layer.

Current progress:

- grid exists
- formula bar exists
- side inspector exists
- normal vs Smart Cells exists
- Runner Preview exists
- Clear Sheet and Load Demo actions exist
- formula reference popup filters and supports keyboard insertion
- Excel-style range formulas and `IF` support exist
- undo/redo exists for sheet edits and structural changes
- formula-aware copy/paste and fill-down exist for coordinate formulas
- row/column add/delete controls exist
- lookup table paste from Excel-style tabular data exists
- inspector dependency tracing exists
- Smart Cell display labels exist
- rule messages are separate from internal annotations
- dropdown runner inputs exist
- action Smart Cell role exists
- multi-criteria lookup editor exists
- Help tab exists
- `.xlsx` import entry point exists

Next work:

- continue Excel-like editing refinements
- improve formula editing/cursor insertion behavior
- improve lookup table ergonomics only where it helps near-term demos; do not treat the current inspector-embedded lookup table editor as the final model
- improve visual polish for Smart Cells, rule cells, and action cells
- refine demo content and runner grouping after user testing

## Phase 2 - Local Configuration Save/Load

Goal: move from one local sheet to named local configurations.

Current progress:

- new configuration
- rename configuration
- save current configuration
- load configuration
- duplicate configuration
- delete local configuration
- Excel import creates a new local configuration
- unsaved-change indicator
- saved sheet dimensions for added/deleted rows and columns

Keep this browser-local first. Do not add the database until the workflow feels right.

Next work:

- make configuration actions more polished after user testing
- consider import/export of local configurations before database work

## Phase 2.25 - Excel Calculator Import

Goal: let existing spreadsheet calculators become Quoin configurations.

Product framing:

```text
Import the calculator, then structure it.
```

Current progress:

- `exceljs` parser dependency added
- neutral import model exists in `lib/import/types.ts`
- workbook reader exists in `lib/import/read.ts`
- imported-sheet converter exists in `lib/import/convert.ts`
- `Import Excel` button exists in the toolbar
- `.xlsx` files can be read in the browser
- user can choose one worksheet from a workbook
- imported values and formulas are preserved in the grid
- imports create new local configurations by default
- safe workbook-defined single-cell names can become Smart Cell names
- named ranges and unsupported/risky formula patterns produce review items
- post-import summary/report exists
- root `import-test-*.xlsx` fixtures exist
- fixture smoke-test script exists
- manual browser fixture results are captured in `Quoin test notes.md`
- private GitHub baseline is established before the next import model changes

Next work:

- design Quoin's workbook structure for multi-sheet imports and reference-table-like data
- update import flow so workbook structure is preserved instead of flattening to one selected worksheet only
- improve import review copy and empty/error states after testing
- decide whether imported named cells should be surfaced automatically or remain unsurfaced by default
- add guided repair affordances for common review items
- consider import of basic Excel data validation lists into runner input options
- consider whether multi-sheet workbooks should become future Quoin workbook tabs or reference tables
- expand formula compatibility based on real imported calculators

## Phase 2.5 - Demo Readiness

Goal: make Quoin easy to show as a replacement direction for spreadsheet-driven shop workflows.

Current progress:

- generic `Demo - Beam Selection` local workflow exists
- demo uses fake/non-proprietary data only
- runner preview groups surfaced inputs, outputs, shop actions, review flags, and validation
- Help tab explains the core concepts

Next work:

- polish the demo data and labels based on feedback
- add a short guided demo script or in-app quick-start
- improve empty/error states for demo flow

## Phase 2.75 - Reference Tables / CSV Lookup Direction

Goal: move lookup behavior toward Excel-familiar reference data instead of hidden per-cell tables.

Product direction:

- CSV ingestion should be supported at minimum.
- Pasted grid ranges should be able to become named reference tables.
- Lookup Smart Cells should be able to query named reference tables.
- Reference tables should feel closer to Excel tabs or named table ranges than cell-local metadata.
- Imported Excel named ranges and table-like regions should inform this model.
- Runner Preview should continue to hide raw reference data and show only surfaced workflow fields/results.

Needed later:

- reference table object model
- CSV import path
- Excel imported range/table promotion path
- named table/range promotion
- lookup Smart Cell binding to reference table, match columns, and output column
- clear missing-match behavior and runner-facing messages
- publish/run snapshot support for table identity, version, and matched row

Keep the current embedded lookup table editor as a prototype shortcut until this model is designed.

## Phase 3 - Publish Snapshot

Goal: separate draft authoring from a published runner version.

Needed:

- draft sheet state
- publish button
- immutable published snapshot
- version number
- Runner Preview can switch between draft preview and published preview

Rule:

- Editing a draft must not silently change a published configuration.

## Phase 4 - Runner Execution Records

Goal: make Runner Preview into a real runner flow.

Needed:

- create run
- enter surfaced inputs
- click Run
- save execution snapshot
- preserve inputs, outputs, warnings, failures, timestamp

Still local-first until the model is proven.

## Phase 5 - Audit Report

Goal: generate a defensible report from an execution.

Needed:

- run ID
- configuration name/version
- user-entered inputs
- calculated outputs
- lookup/rule results
- validation failures
- compliance warnings
- timestamp/build identifier
- printable HTML report first

PDF export can come after the report content is right.

## Later Platform Phases

Deferred:

- real database
- auth
- accounts/departments/users
- invite flow
- project/run hierarchy
- AI ingestion
- published outputs and output binding
- full audit/export system
- SaaS deployment
