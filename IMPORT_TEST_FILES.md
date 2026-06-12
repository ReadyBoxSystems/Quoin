# IMPORT_TEST_FILES.md

These `.xlsx` files are root-level fixtures for manually testing Quoin's Excel importer.

## Files

- `import-test-01-basic-values.xlsx`
  - Simple labels, numbers, text, boolean, and date values.

- `import-test-02-basic-formulas.xlsx`
  - Basic formulas using multiplication, `ROUND`, and `IF`.

- `import-test-03-named-cells.xlsx`
  - Workbook-defined names for safe Smart Cell promotion, plus an unsafe name and a named range that should become review items.

- `import-test-04-multi-sheet-cross-sheet.xlsx`
  - Two worksheets. The `Calculator` sheet has formulas that reference the `Inputs` sheet and should produce formula review warnings.

- `import-test-05-reference-table-style.xlsx`
  - A visible table-like range with a named range, meant to exercise future reference-table direction.

- `import-test-06-review-warnings.xlsx`
  - Formulas intentionally designed to trigger importer review warnings for cross-sheet references, structured table references, external workbook references, semicolon separators, and spill markers.

## Regenerate

```bash
npm run fixtures:generate
```

## Smoke Test

```bash
npm run fixtures:smoke
```

