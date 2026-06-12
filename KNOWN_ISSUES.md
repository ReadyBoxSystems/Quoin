# KNOWN_ISSUES.md

## Active Issues / Rough Edges

- Formula reference popup is functional but basic.
- Formula editing still needs closer Excel-like behavior.
- Lookup editor supports multi-criteria lookup and Excel-style paste, but large-table browsing/searching is still rough.
- Lookup cell display/error behavior is prototype-level.
- Lookup tables currently live inside Smart Cell metadata. This is useful for the prototype but is probably not the final model; the intended direction is first-class reference tables / CSV-ingested datasets that lookup Smart Cells query.
- Excel import exists, but it is intentionally not full Excel compatibility.
- Excel import reads one worksheet at a time and does not yet preserve a whole workbook as multiple Quoin tabs.
- Imported cross-sheet, external workbook, structured table, spill/array, and semicolon-separator formulas are preserved but flagged for review.
- Imported review items are informational only; there is not yet a guided repair workflow.
- Imported Excel formatting, charts, pivots, macros/VBA, merged-cell behavior, and table semantics are not supported.
- Dropdown option editing is basic and lives in a newline/comma-separated inspector field.
- Row/column add/delete exists, but there is no right-click context menu or drag-based row/column resizing yet.
- Formula-aware copy/fill adjusts coordinate references, but there is no mouse drag fill handle yet.
- Dependency tracing is inspector-only; it does not yet draw arrows or highlight precedents/dependents in the grid.
- Local storage can contain stale sheet/configuration state across major prototype changes. Use `Load Demo`, `Clear Sheet`, or create a new configuration if the demo looks wrong.
- The app has no database, so browser/local storage is the only persistence.
- Local named configurations exist, but there is no publish snapshot yet.
- There are no execution records or audit reports yet.
- The generic beam demo uses fake data only and is not an engineering calculator.

## Dev Server Notes

Avoid running `npm run build` while the dev server is open and the user is testing. It has previously caused stale `.next` chunk errors.

If weird Next.js runtime chunk errors appear:

1. Stop the dev server.
2. Delete generated `.next`.
3. Restart with `npm run dev` or `Quoin.bat`.

## Current Prototype Limitations

- Runner Preview is generated live from the draft sheet.
- Validation failure marks the run invalid but does not create a saved failed execution.
- Compliance warning appears but is not yet tied to an execution record.
- Surface toggle controls Runner Preview visibility only.
- Lookup runs do not yet record reference table identity, table version, or matched source row.
- Local configuration changes are saved only when the user clicks Save.
- Excel workbook import exists, but there is no Quoin-native configuration import/export yet.
- There is no permission model yet.

## Import Fixture Notes

Root-level `import-test-*.xlsx` files are generated test fixtures for manual importer testing.

Use:

```bash
npm run fixtures:generate
npm run fixtures:smoke
```

The fixture smoke test checks the importer pipeline, but browser UI testing is still needed.
