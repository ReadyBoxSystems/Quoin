# STATUS.md

## Current State

Quoin is now a clean build in this folder.

The project is now tracked in the private GitHub repository:

```text
https://github.com/ReadyBoxSystems/Quoin.git
```

The local working branch is `main`, tracking `origin/main`. The first pushed baseline is commit `5259f17` (`Initial Quoin prototype baseline`).

Current Git state as of July 17, 2026:

- `origin/main` and local `main` are aligned at `b5162ef` (`Archive import fixtures and update planning docs`).
- The preceding pushed commit is `2dcded4` (`Support lookup formulas as calculation primitives`).
- Completed importer planning notes, manual test notes, and generated workbook fixtures are archived under `docs/archive/import-testing/`.

The prototype has the right core shape:

- spreadsheet-style grid with column letters and row numbers
- compact spreadsheet grid sizing with a larger default visible sheet area
- normal coordinate cells
- Smart Cell promotion by naming a cell
- side inspector for Smart Cell metadata
- formula bar
- Sheet tabs directly below the formula bar
- Sheet rename support that updates direct cross-Sheet references using the old Sheet name
- sticky row and column headers while scrolling the grid
- formula references by coordinate or Smart Cell name
- workbook-scoped Smart Cell name references across Sheets
- Excel-style range formulas such as `SUM(A1:A5)` and rectangular ranges such as `SUM(A1:B3)`
- `IF(...)` formulas and common uppercase Excel function aliases
- `ROUNDUP(...)`, exact-match `VLOOKUP(...)`, exact/default `XLOOKUP(...)`, cross-Sheet lookup ranges, and common helper-key lookup formulas
- formula reference popup with filtering and keyboard insertion
- mouse selection/editing cleaned up so click selection and double-click editing do not accidentally clear cell values
- undo/redo for sheet edits
- formula-aware copy/paste and fill-down for coordinate formulas
- dynamic add/delete row and column controls
- dependency tracing in the inspector
- Smart Cell display labels separate from formula-safe names
- lookup, action, validation, and compliance cell roles
- multi-criteria lookup editor
- tab/newline paste from Excel into lookup tables
- dropdown inputs in the grid and Runner Preview
- imported Excel dropdown inputs promoted to surfaced Smart Cells when supported
- Runner Preview generated from workbook-wide surfaced Smart Cells with Sheet-aware input/output/action/review/validation grouping
- Runner Preview hides empty optional sections
- local browser persistence with named configurations
- `.xlsx` import into new local configurations with preserved Sheets
- importer review items for merged ranges, named ranges, unsafe names, and likely formula compatibility issues
- importer review items for unsupported Excel dropdown sources
- importer review items for approximate lookup behavior, while exact lookup formulas remain normal formula cells
- grid columns auto-size to content within bounded widths
- expanded Help tab with workflow, Smart Cell, dropdown, formula, import, and keyboard guidance
- generic fake-data `Demo - Beam Selection` workflow
- double-click launcher via `Quoin.bat`
- private GitHub repository initialized and pushed

## What Works

- Values and formulas can be entered into grid cells.
- Compact default grid shows more rows and columns without making cells visually oversized.
- Named Smart Cells can be referenced by formula.
- Unique Smart Cell names can be referenced directly by formulas from any Sheet in the workbook.
- Coordinate references work.
- Formula ranges, rectangular ranges, `IF`, `ROUNDUP`, common Excel-style function aliases, exact-match `VLOOKUP`, exact/default `XLOOKUP`, lookup formulas inside larger formulas, cross-Sheet lookup ranges, and concatenated helper-key lookups work.
- Copy/paste and fill-down adjust coordinate references while leaving Smart Cell names alone.
- Rows and columns can be added or deleted; formulas shift coordinate references and deleted references become `#REF!`.
- Undo/redo covers sheet edits, metadata changes, paste, clear, demo reload, runner input changes, and row/column structure changes.
- The inspector shows upstream dependencies and downstream dependents for the selected cell.
- Lookup cells can use editable multi-criteria lookup tables.
- Lookup tables accept paste from Excel-style tabular data.
- Action cells can surface shop notes or required runner actions.
- Smart Cells with dropdown options render as dropdowns in the grid and Runner Preview.
- Imported Excel dropdown cells with literal list sources, same-Sheet ranges, cross-Sheet ranges, quoted Sheet ranges, or workbook-defined named ranges become surfaced input Smart Cells with embedded dropdown options.
- Unsupported Excel dropdown sources remain visible as cells and produce review items.
- Compliance cells display `OK` or `WARN`.
- Validation cells display `PASS` or `FAIL`.
- Runner Preview shows workbook-wide surfaced inputs, outputs, shop actions, review flags, validation, and rule messages, grouped by Sheet when needed.
- Local configurations can be created, renamed, saved, loaded, duplicated, and deleted.
- Excel workbooks can be imported as multi-Sheet local configurations.
- Imported values and formulas are preserved in the Quoin grid.
- Imported supported Excel data-validation dropdowns are preserved as Quoin dropdown input controls.
- Imported formulas remain visible when unsupported or risky features need review.
- Imported exact lookup formulas remain normal formula cells and calculate without being converted into lookup Smart Cells.
- Approximate `VLOOKUP` and non-exact `XLOOKUP` behavior is preserved for review.
- Safe workbook-defined names pointing to single cells can promote imported cells to Smart Cells.
- Supported imported dropdown cells without workbook-defined names are given formula-safe generated Smart Cell names and labels based on nearby row labels when possible.
- Imported merged ranges, named ranges, unsafe names, external workbook references, structured table references, spill markers, and semicolon separators are surfaced as review items.
- Clear Sheet blanks the workspace; Load Demo restores the generic demo sheet.
- Historical import fixture workbooks are archived under `docs/archive/import-testing/`.
- Engine and import tests pass.

## Prototype-Only Areas

- State is stored in browser local storage, not a database.
- Local configurations are browser-local only.
- Lookup editor supports Excel paste but is still not a full large-table management surface.
- Excel import is a first local prototype, not full Excel compatibility.
- Imported dropdown source ranges are snapshotted into embedded options, not live-linked to the source range.
- Import now preserves workbook sheets as Quoin Sheets, and Runner Preview evaluates surfaced Smart Cells across the workbook.
- Basic imported cross-sheet formulas, workbook-scoped Smart Cell name references, exact lookup formulas, and cross-Sheet lookup ranges calculate across preserved Sheets; more advanced workbook features are still preserved and flagged for review.
- Rule messages are separate from annotations, but there is no rule library yet.
- No publish snapshot yet.
- No immutable execution records yet.
- No audit report yet.
- No auth, account, department, or user model yet.

## Recent Decisions

- Quoin should feel like Excel first, then add Smart Cell behavior.
- Excel power-user expectations are now part of demo readiness: basic range formulas, `IF`, undo/redo, formula fill, lookup paste, resizing, and dependency tracing should not regress.
- Admin workflow should be:
  1. Build the spreadsheet.
  2. Prove the math works.
  3. Name important cells.
  4. Surface runner-facing cells.
  5. Let Runner Preview show the controlled form.
- Validation means run failure, not stopped math.
- Compliance means warning, not invalid result.
- Smart Cell Name should stay formula/backend-safe; Display Label is what runners see.
- Demo data must remain fake/non-proprietary.
- Quoin should improve the workflow direction, not visually clone an existing workbook.
- Excel import should mean "import the calculator, then structure it"; Quoin should preserve the calculation surface first and keep Smart Cell mapping as a deliberate follow-up step.
- Unsupported Excel features should produce review items rather than silently dropping formulas or workbook names.
- The visible UI label for workbook tabs is "Sheet"; "Workbook Structure" describes the underlying model.
- Sheet navigation belongs near the formula bar, directly above the grid, not in the side inspector area.
- What admins configure should be visible where the cell is used: dropdown Smart Cells appear as dropdowns in the grid and Runner Preview.
- Short embedded dropdowns are acceptable now; live reference-data-backed dropdowns belong with the later reference table model.
- Lookup Smart Cells are for Quoin-native structure, auditability, runner surfacing, and editable criteria. They are not required for normal exact lookup formula compatibility.

## Verification

Last known verification commands:

```bash
npm.cmd run typecheck
npm.cmd run test:engine
npm.cmd run test:import
npm.cmd run test:import-reader
```

These were passing after the latest changes.

Manual testing still needed:

- Define the next manual testing approach before adding or regenerating workbook fixtures.
- If import behavior changes again, create focused fixtures for that specific behavior rather than restoring the old broad fixture sweep.
- Manually test visible Help copy after the app copy is made brand-neutral.

## Open UX Notes

- Formula popup is useful and filterable but still needs closer Excel-like cursor behavior.
- Keyboard navigation was fixed so focus follows selected cell.
- Failed lookup cells show `#ERR`; messages include failed criteria.
- Multi-criteria lookup works and supports pasted tabular data, but large-table browsing/searching still needs polish.
- Exact lookup formulas work in normal cells; lookup Smart Cells remain a separate Quoin-native structure.
- Dropdown inputs work in the grid and Runner Preview, but admin editing of dropdown option lists is basic.
- Imported dropdown inputs now work in the grid, inspector, and Runner Preview when their Excel data-validation source is supported.
- The grid/inspector relationship feels directionally right but will need more polish.
- Import review output is useful but basic; it lists issues before confirmation but does not yet provide guided repair actions.
- The import flow is functional but not yet a polished wizard.

## Lookup Table Direction

The current embedded lookup table editor is a prototype shortcut, not necessarily the final product model.

Preferred direction:

- treat lookup data as first-class reference tables or datasets
- support CSV ingestion at minimum
- allow pasted grid/range data to become named reference tables
- use imported named ranges and visible workbook table-like areas as signals for future reference tables
- let lookup Smart Cells query those named reference tables
- keep Runner Preview focused on surfaced inputs, outputs, actions, warnings, and validation
- eventually record table identity/version and matched row in publish/run/audit flows

This is intentionally not an implementation commitment for the immediate demo. The near-term rule is to avoid hardening the embedded per-cell lookup editor in a way that blocks a future reference-table model.
