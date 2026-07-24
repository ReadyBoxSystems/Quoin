# Quoin

Quoin is a spreadsheet-first tool for turning shop knowledge into structured, repeatable workflows.

Many teams run on a mix of spreadsheet calculators, reference PDFs, tribal knowledge, and handwritten judgment calls. That works until a calculator is hard to audit, the expert is unavailable, or a new hire has to guess which assumptions matter.

Quoin starts where those teams already work: a familiar grid. Admins can build or import calculator logic, prove the math, then promote the important cells into Smart Cells that power a controlled Runner Preview.

## What Quoin Does

- Provides an Excel-familiar authoring grid for normal values and formulas.
- Supports common spreadsheet behavior such as coordinate references, ranges, `IF`, `COUNT`, `ROUNDUP`, exact-match `VLOOKUP`, and exact/default `XLOOKUP`.
- Preserves multi-Sheet workbook structure during local `.xlsx` import.
- Lets admins name important cells, turning them into Smart Cells with roles, labels, annotations, dropdowns, lookup behavior, actions, validation, and compliance warnings.
- Generates Runner Preview from surfaced Smart Cells so the runner sees the controlled workflow instead of the calculator guts.
- Stores configurations locally in the browser while the product model is still being proven.

## Workflow

1. Build a calculator directly in Quoin, or import an existing workbook.
2. Verify the math in the spreadsheet grid.
3. Name the important cells to promote them into Smart Cells.
4. Add display labels, roles, dropdown options, rules, lookup behavior, and runner surfacing.
5. Use Runner Preview to see the controlled workflow generated from the Smart Cells.
6. Save the configuration locally in the browser.

The goal is not to force teams to design an app before doing spreadsheet work. The grid comes first. The structure is added where it helps.

## Current Prototype Scope

Quoin is an internal prototype under active development.

Current scope:

- Next.js App Router
- React and TypeScript
- deterministic formula engine powered by `mathjs`
- local browser storage
- local named configurations
- local `.xlsx` workbook import
- no backend database
- no authentication
- no published execution records yet

The included demo uses fake, non-proprietary data and is meant to show the workflow direction, not provide engineering guidance.

## Running Locally

On Windows, double-click:

```bat
Quoin.bat
```

The launcher starts the development server and opens:

```text
http://localhost:3000
```

Or run it from a terminal:

```bash
npm install
npm run dev
```

Then open:

```text
http://localhost:3000
```

## Status

This repository contains the runnable prototype app. Internal planning notes, local test files, generated fixtures, logs, and reference documents are intentionally kept out of the GitHub-tracked app surface.
