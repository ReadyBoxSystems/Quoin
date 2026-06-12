# AGENTS.md

## Project

Quoin is being rebuilt cleanly in this folder as a "better Excel" for structured shop knowledge.

The product direction is:

- Start with an Excel-familiar grid.
- Normal coordinate cells should work first.
- Excel-familiar behavior matters for demos: range formulas, `IF`, undo/redo, formula-aware copy/fill, paste-friendly lookup tables, row/column controls, and dependency tracing should keep working.
- Naming a cell promotes it to a Smart Cell.
- Smart Cells can have role, type, display label, annotation, runner surfacing, lookup behavior, action output behavior, and rule behavior.
- Runner Preview is generated from surfaced Smart Cells.
- Local named configurations are browser-local for now.
- Excel import is now an early local prototype path: import an `.xlsx` calculator, choose one worksheet, preserve values/formulas, create a new local configuration, and then structure it with Quoin Smart Cells.
- Lookup table direction: keep the current embedded Smart Cell lookup tables as a prototype shortcut, but the intended product model is first-class reference tables / CSV-ingested datasets that lookup Smart Cells can query. This should feel closer to Excel tabs or named table ranges than hidden per-cell data.
- AI ingestion, database persistence, auth, departments, publishing, and audit reports are later phases.

Reference docs are stored in `docs/reference/`.

## How To Run

The preferred user entry point is:

```bat
Quoin.bat
```

Double-clicking `Quoin.bat` starts the Next.js dev server and opens:

```text
http://localhost:3000
```

Developer commands:

```bash
npm run typecheck
npm run test:engine
npm run test:import
npm run test:import-reader
npm run fixtures:smoke
```

Avoid running `npm run build` while the user is actively testing the dev server. Next.js production builds can disturb the running `.next` dev output. If a production build is needed, stop the dev server first.

## Current Architecture

This is currently a local prototype with:

- Next.js App Router
- TypeScript
- React client-side Variable Sheet
- `mathjs`-based deterministic engine
- Excel-style formula support for coordinate references, ranges, common function aliases, and `IF`
- browser local storage persistence
- browser-local named configurations
- ExcelJS-based `.xlsx` workbook reader isolated behind `lib/import/read.ts`
- neutral import model and Quoin sheet converter in `lib/import/`
- no database yet
- no auth yet

Important files:

- `app/page.tsx` - renders the prototype shell
- `components/variable-sheet.tsx` - main sheet, inspector, runner preview
- `lib/engine/index.ts` - calculation engine
- `lib/engine/types.ts` - engine types
- `lib/engine/engine.test.ts` - engine tests
- `lib/import/types.ts` - neutral Excel import model
- `lib/import/read.ts` - `.xlsx` workbook reader
- `lib/import/convert.ts` - imported sheet to Quoin grid converter
- `IMPORT_PLAN.md` - Excel import plan and scope
- `IMPORT_TEST_FILES.md` - root fixture guide for importer testing
- `Quoin.bat` - double-click launcher

## Product Rules

- Do not turn Quoin into a card/dashboard app. The authoring surface should remain spreadsheet-first.
- Do not force users to design an app before doing spreadsheet work.
- A normal cell should feel like a normal spreadsheet cell.
- Excel import should bring in the calculator grid first; Smart Cell roles, labels, surfacing, and runner behavior remain Quoin-native follow-up work.
- Preserve Excel-like muscle memory unless it conflicts with the Smart Cell model.
- A Smart Cell is created by naming a cell.
- Smart Cell metadata belongs in the side inspector, not crammed into the grid.
- Smart Cell names are formula/backend-safe. Use Display Label for human-facing runner text.
- Surfaced Smart Cells feed Runner Preview.
- Normal coordinate cells do not appear in Runner Preview.
- Action Smart Cells are for runner-facing shop actions/notes.
- Validation is PASS/FAIL, not "block execution." Math should still run where possible.
- Compliance is OK/WARN. Warnings do not invalidate the calculation.
- Incomplete formulas should not crash the app while the user is typing.
- Imported formulas should remain visible even if Quoin cannot calculate them yet. Unsupported or risky Excel features should produce review items, not silent data loss.
- Long-term lookup behavior should center on visible/imported reference data: pasted grids, named ranges/tables, and CSV imports that can be versioned and audited later. Do not over-invest in the current inspector-embedded lookup table UI as the final model.

## Development Rules

- Keep the engine independent from React, Next.js, and persistence.
- Add engine tests for calculation behavior changes.
- Add import tests for workbook parsing, conversion, and formula review changes.
- During active UI iteration, run `npm run typecheck` and `npm run test:engine`.
- During Excel import work, also run `npm run test:import`, `npm run test:import-reader`, and `npm run fixtures:smoke`.
- Formula calculation changes need focused engine tests.
- Prefer small UX iterations that preserve the spreadsheet mental model.
- Do not modify or rely on the old `C:\Quoin` project unless the user explicitly asks.

## Next Likely Work

See `ROADMAP.md`.
