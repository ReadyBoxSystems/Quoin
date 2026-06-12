# STATUS.md

## Current State

Quoin is now a clean build in this folder.

The prototype has the right core shape:

- spreadsheet-style grid with column letters and row numbers
- normal coordinate cells
- Smart Cell promotion by naming a cell
- side inspector for Smart Cell metadata
- formula bar
- formula references by coordinate or Smart Cell name
- Excel-style range formulas such as `SUM(A1:A5)` and rectangular ranges such as `SUM(A1:B3)`
- `IF(...)` formulas and common uppercase Excel function aliases
- formula reference popup with filtering and keyboard insertion
- undo/redo for sheet edits
- formula-aware copy/paste and fill-down for coordinate formulas
- dynamic add/delete row and column controls
- dependency tracing in the inspector
- Smart Cell display labels separate from formula-safe names
- lookup, action, validation, and compliance cell roles
- multi-criteria lookup editor
- tab/newline paste from Excel into lookup tables
- dropdown runner inputs
- Runner Preview generated from surfaced Smart Cells with generic input/output/action/review/validation grouping
- local browser persistence with named configurations
- `.xlsx` import into new local configurations
- importer review items for named ranges, unsafe names, and likely formula compatibility issues
- Help tab
- generic fake-data `Demo - Beam Selection` workflow
- double-click launcher via `Quoin.bat`

## What Works

- Values and formulas can be entered into grid cells.
- Named Smart Cells can be referenced by formula.
- Coordinate references work.
- Formula ranges, rectangular ranges, `IF`, and common Excel-style function aliases work.
- Copy/paste and fill-down adjust coordinate references while leaving Smart Cell names alone.
- Rows and columns can be added or deleted; formulas shift coordinate references and deleted references become `#REF!`.
- Undo/redo covers sheet edits, metadata changes, paste, clear, demo reload, runner input changes, and row/column structure changes.
- The inspector shows upstream dependencies and downstream dependents for the selected cell.
- Lookup cells can use editable multi-criteria lookup tables.
- Lookup tables accept paste from Excel-style tabular data.
- Action cells can surface shop notes or required runner actions.
- Runner inputs can be free text or controlled dropdowns.
- Compliance cells display `OK` or `WARN`.
- Validation cells display `PASS` or `FAIL`.
- Runner Preview shows surfaced inputs, outputs, shop actions, review flags, validation, and rule messages.
- Local configurations can be created, renamed, saved, loaded, duplicated, and deleted.
- Excel workbooks can be imported from `.xlsx` files one worksheet at a time.
- Imported values and formulas are preserved in the Quoin grid.
- Imported formulas remain visible when unsupported or risky features need review.
- Safe workbook-defined names pointing to single cells can promote imported cells to Smart Cells.
- Imported named ranges, unsafe names, cross-sheet references, external workbook references, structured table references, spill markers, and semicolon separators are surfaced as review items.
- Clear Sheet blanks the workspace; Load Demo restores the generic demo sheet.
- Import fixture workbooks exist at the repo root for manual importer testing.
- Engine and import tests pass.

## Prototype-Only Areas

- State is stored in browser local storage, not a database.
- Local configurations are browser-local only.
- Lookup editor supports Excel paste but is still not a full large-table management surface.
- Excel import is a first local prototype, not full Excel compatibility.
- Import supports one worksheet at a time and does not yet preserve full workbook structure as Quoin tabs/reference tables.
- Imported cross-sheet formulas are preserved but not fully supported by the engine.
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

## Verification

Last known verification commands:

```bash
npm run typecheck
npm run test:engine
npm run test:import
npm run test:import-reader
npm run fixtures:smoke
```

These were passing after the latest changes.

Manual testing still needed:

- Run the root `import-test-*.xlsx` files through the browser import UI.
- Confirm the import summaries and review items match `IMPORT_TEST_FILES.md`.
- Check that each import creates a new local configuration and does not overwrite existing work.

## Open UX Notes

- Formula popup is useful and filterable but still needs closer Excel-like cursor behavior.
- Keyboard navigation was fixed so focus follows selected cell.
- Failed lookup cells show `#ERR`; messages include failed criteria.
- Multi-criteria lookup works and supports pasted tabular data, but large-table browsing/searching still needs polish.
- Dropdown inputs work in Runner Preview, but admin editing of dropdown option lists is basic.
- The grid/inspector relationship feels directionally right but will need more polish.
- Import review output is useful but basic; it lists issues but does not yet provide guided repair actions.
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
