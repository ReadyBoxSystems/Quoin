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
- Preserve workbook Sheets and let the user choose which Sheet opens first.
- Preserve normal cell values.
- Preserve formulas as formulas.
- Keep unsupported formulas visible in the grid instead of dropping them.
- Expand Quoin row/column dimensions to fit imported content.
- Create a new local Quoin configuration by default.
- Run Quoin's engine after import so formula errors and unsupported behavior are visible.

Current status: the first local implementation covers this target.

Optional if straightforward:

- Import workbook-defined names that point to a single cell as Smart Cell names. Done for safe formula/backend-safe names.
- Import supported Excel data-validation list dropdowns as Quoin dropdown inputs. Done for literal lists, same-workbook ranges, quoted Sheet ranges, and workbook-defined named ranges.
- Detect names that point to ranges and report them for review. Done as review items.
- Detect merged ranges and report them for review without duplicating covered-cell values. Done as review items.
- Preserve basic row/column dimensions.

Out of scope for this first importer:

- Full Excel visual formatting.
- Charts.
- Pivot tables.
- Macros/VBA.
- External workbook links.
- Perfect merged-cell visual rendering.
- Full table semantics.
- Full Excel workbook formula semantics beyond basic cross-sheet references.
- Automatic Smart Cell role inference beyond clear named-cell cases.
- Live reference-data-backed dropdown binding.
- Publishing, execution records, or audit reporting.

## Product Rules

- Imported normal cells should behave like normal Quoin grid cells.
- Import should not force users to design the runner flow first.
- Smart Cell metadata remains Quoin-native and belongs in the inspector.
- Unsupported Excel features should be surfaced as review items, not silently discarded.
- Existing formulas should remain visible even if Quoin cannot calculate them yet.
- Importing should not overwrite the active configuration without explicit confirmation.
- Supported Excel dropdown inputs should become visible Quoin dropdown controls immediately after import.

## Suggested User Flow

1. Admin clicks `Import Excel`.
2. Admin selects an `.xlsx` file.
3. Quoin reads the workbook and lists Sheets.
4. Admin chooses which Sheet opens first.
5. Quoin shows a short import review before confirmation:
   - selected Sheet
   - cells imported
   - formulas imported
   - named cells/ranges found
   - merged ranges, formulas, or references needing review
6. Admin confirms import.
7. Quoin creates a new local configuration from the imported workbook Sheets.
8. Admin reviews the grid, names important cells, sets roles, and surfaces runner-facing Smart Cells.

## Technical Shape

Keep the importer separate from the engine and React UI.

Recommended layers:

- Workbook reader: parses `.xlsx` into a neutral import model.
- Import model: describes sheets, cells, formulas, and workbook names without React state.
- Quoin converter: converts imported Sheets into Quoin grid cells and dimensions. This lives in `lib/import/convert.ts`.
- Dropdown translation: the reader captures Excel list validations, and the converter applies options/promotes supported dropdown cells.
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
  merges?: ImportedMerge[];
  rowCount: number;
  columnCount: number;
};

type ImportedCell = {
  address: string;
  value?: string | number | boolean | null;
  formula?: string;
};

type ImportedDataValidation = {
  address: string;
  type: "list";
  options?: string[];
  source?: string;
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

- external workbook references
- structured table references
- dynamic array or implicit-intersection markers
- semicolon argument separators

Likely early unsupported areas:

- structured table references
- external workbook references
- advanced Excel functions not yet implemented by the Quoin engine
- array/spill formulas

Basic cross-sheet cell references across imported Sheets are now supported by the workbook execution path.
Unique Smart Cell names are workbook-scoped, so formulas on any Sheet can reference them directly by name.

## Named Cells And Ranges

If the workbook defines names:

- Names pointing to a single cell may become Smart Cell names, if formula/backend-safe.
- Unsafe names should be sanitized or reported for review before promotion.
- Names pointing to ranges should not become Smart Cells automatically.
- Named ranges may later become candidates for first-class reference tables.

This should be conservative. The admin can always name cells manually after import.

## Dropdown Import

Excel data-validation list dropdowns now translate into Quoin input dropdowns when the source can be safely snapshotted.

Supported source shapes:

- typed literal lists such as `"standard,heavy"`
- same-Sheet ranges such as `=$A$2:$A$8`
- cross-Sheet ranges such as `=Lists!$A$2:$A$8`
- quoted Sheet names such as `='Reference Data'!$B$2:$B$5`
- workbook-defined names that resolve to one range

Import behavior:

- snapshot current source values into embedded `inputOptions`
- skip blank source values
- dedupe duplicate values in first-seen order
- promote supported dropdown cells to surfaced input Smart Cells
- use workbook-defined names when present, otherwise generate safe names from nearby row labels when possible
- leave unsupported table/formula/external dropdown sources as review items

This is intentionally not a live reference-data binding. Live dropdown sources belong with the future reference table model.

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
- choose which Sheet opens first - done
- import workbook Sheets, values, and formulas into a new local configuration - done
- preserve formulas visibly - done
- detect merged ranges without duplicating covered-cell values - done
- expand grid dimensions - done
- show an import review before confirmation - done
- leave most Smart Cell mapping to the existing inspector workflow - done
- promote supported imported dropdown inputs so the imported runner-facing controls are visible immediately - done

Manual browser testing has started against real-world calculators and generated root fixtures.

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
- Improve import review wording and guided repair actions based on testing.
- Decide whether imported single-cell names should stay unsurfaced by default.
- Improve unsupported Excel dropdown source repair affordances.
- Use real calculator imports to prioritize additional Excel functions and formula syntax support.
- Design how imported named ranges/table-like areas should become first-class Quoin reference tables.
