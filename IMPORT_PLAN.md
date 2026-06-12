# IMPORT_PLAN.md

## Goal

Build Quoin's first Excel import path so existing spreadsheet calculators can become Quoin configurations.

The product framing is:

```text
Import the calculator, then structure it.
```

This phase should import the spreadsheet calculation surface, not attempt to clone all of Excel.

## Phase Scope

Initial import target:

- `.xlsx` files only.
- Read workbook sheet names.
- Import one worksheet at a time.
- Preserve normal cell values.
- Preserve formulas as formulas.
- Keep unsupported formulas visible in the grid instead of dropping them.
- Expand Quoin row/column dimensions to fit imported content.
- Create a new local Quoin configuration by default.
- Run Quoin's engine after import so formula errors and unsupported behavior are visible.

Current status: the first local implementation covers this target.

Optional if straightforward:

- Import workbook-defined names that point to a single cell as Smart Cell names. Done for safe formula/backend-safe names.
- Detect names that point to ranges and report them for review. Done as review items.
- Preserve basic row/column dimensions.

Out of scope for this first importer:

- Full Excel visual formatting.
- Charts.
- Pivot tables.
- Macros/VBA.
- External workbook links.
- Perfect merged-cell behavior.
- Full table semantics.
- Full cross-sheet formula support.
- Automatic Smart Cell role inference beyond clear named-cell cases.
- Publishing, execution records, or audit reporting.

## Product Rules

- Imported normal cells should behave like normal Quoin grid cells.
- Import should not force users to design the runner flow first.
- Smart Cell metadata remains Quoin-native and belongs in the inspector.
- Unsupported Excel features should be surfaced as review items, not silently discarded.
- Existing formulas should remain visible even if Quoin cannot calculate them yet.
- Importing should not overwrite the active configuration without explicit confirmation.

## Suggested User Flow

1. Admin clicks `Import Excel`.
2. Admin selects an `.xlsx` file.
3. Quoin reads the workbook and lists worksheets.
4. Admin chooses one worksheet.
5. Quoin shows a short import summary:
   - selected worksheet
   - cells imported
   - formulas imported
   - named cells/ranges found
   - formulas or references needing review
6. Admin confirms import.
7. Quoin creates a new local configuration from the imported sheet.
8. Admin reviews the grid, names important cells, sets roles, and surfaces runner-facing Smart Cells.

## Technical Shape

Keep the importer separate from the engine and React UI.

Recommended layers:

- Workbook reader: parses `.xlsx` into a neutral import model.
- Import model: describes sheets, cells, formulas, and workbook names without React state.
- Quoin converter: converts one imported sheet into Quoin grid cells and dimensions. This lives in `lib/import/convert.ts`.
- UI flow: file picker, sheet selector, import summary, and confirmation. The first local version is wired into `components/variable-sheet.tsx`.

The engine should continue to receive normal Quoin cells only. It should not know about Excel file parsing.

Implemented files:

- `lib/import/types.ts`
- `lib/import/read.ts`
- `lib/import/convert.ts`
- `components/variable-sheet.tsx`
- `scripts/generate-import-fixtures.mjs`
- `scripts/smoke-import-fixtures.mjs`
- `IMPORT_TEST_FILES.md`

## Parser Decision

Use `exceljs` for the first importer.

Reasoning:

- It reads `.xlsx` workbooks in JavaScript.
- It exposes worksheets and cell addresses.
- It preserves formula cells so Quoin can keep formulas visible.
- It can support workbook metadata such as defined names.
- It works in the same TypeScript/Next.js environment as the current prototype.

Package note:

- `xlsx` / SheetJS was considered first because it is widely used, but the npm package currently reports unresolved high-severity audit advisories.
- `exceljs` currently avoids those high-severity `xlsx` advisories, but it does bring a moderate transitive `uuid` audit advisory and some older transitive packages.
- Keep parser usage isolated behind the workbook reader so Quoin can swap libraries later if needed.

## Neutral Import Model

The importer should convert Excel data into a simple internal structure before touching Quoin state.

The neutral model lives in `lib/import/types.ts`.

Example shape:

```ts
type ImportedWorkbook = {
  sheets: ImportedSheet[];
  names: ImportedName[];
};

type ImportedSheet = {
  name: string;
  cells: ImportedCell[];
  rowCount: number;
  columnCount: number;
};

type ImportedCell = {
  address: string;
  value?: string | number | boolean | null;
  formula?: string;
};

type ImportedName = {
  name: string;
  sheetName?: string;
  reference: string;
  kind: "singleCell" | "range" | "other";
};
```

## Formula Handling

Formula import should preserve the source formula text as much as possible.

Rules:

- Excel formulas should become Quoin entries starting with `=`.
- Supported formulas should calculate normally.
- Unsupported formulas should remain in the cell and show an error/review state.
- Unsupported formulas should help build the future engine backlog.
- Formula review should be specific enough to tell the admin what needs attention.

Current first-pass review flags:

- cross-sheet references
- external workbook references
- structured table references
- dynamic array or implicit-intersection markers
- semicolon argument separators

Likely early unsupported areas:

- cross-sheet references
- structured table references
- external workbook references
- advanced Excel functions not yet implemented by the Quoin engine
- array/spill formulas

## Named Cells And Ranges

If the workbook defines names:

- Names pointing to a single cell may become Smart Cell names, if formula/backend-safe.
- Unsafe names should be sanitized or reported for review before promotion.
- Names pointing to ranges should not become Smart Cells automatically.
- Named ranges may later become candidates for first-class reference tables.

This should be conservative. The admin can always name cells manually after import.

## Reference Table Direction

Excel named ranges, tables, and pasted data ranges are important signals for Quoin's future reference-table model.

For this first importer:

- Preserve the grid data.
- Report named ranges/tables if detectable.
- Do not overbuild table management into the importer yet.

Later, imported ranges should be promotable to visible Quoin reference tables or CSV-backed datasets.

## First Milestone

The first completed build should demonstrate:

- choose `.xlsx` - done
- choose worksheet - done
- import values and formulas into a new local configuration - done
- preserve formulas visibly - done
- expand grid dimensions - done
- show an import summary - done
- leave Smart Cell mapping to the existing inspector workflow - done

Manual browser testing is still pending against the generated root fixtures.

Verification:

```bash
npm run typecheck
npm run test:engine
npm run test:import
npm run test:import-reader
npm run fixtures:smoke
```

## Fixture Workbooks

Generated root fixtures:

- `import-test-01-basic-values.xlsx`
- `import-test-02-basic-formulas.xlsx`
- `import-test-03-named-cells.xlsx`
- `import-test-04-multi-sheet-cross-sheet.xlsx`
- `import-test-05-reference-table-style.xlsx`
- `import-test-06-review-warnings.xlsx`

See `IMPORT_TEST_FILES.md` for the manual test sequence and expected behavior.

## Next Import Work

- Run manual browser tests with each generated fixture.
- Improve import review panel wording and empty/error states based on testing.
- Decide whether imported single-cell names should stay unsurfaced by default.
- Consider importing Excel data validation lists into Smart Cell input options.
- Use real calculator imports to prioritize additional Excel functions and formula syntax support.
- Design how imported named ranges/table-like areas should become first-class Quoin reference tables.
