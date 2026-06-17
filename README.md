# Quoin

Most shops run on a combination of PDF references, Excel calculators, and knowledge that lives in people's heads. That works until it doesn't — when the calculator is wrong, when the expert is out, when the new hire guesses. Quoin is a tool for capturing that logic in a structured, auditable form and surfacing it as a controlled workflow for whoever needs it.

## How It Works

Quoin starts with an Excel-familiar spreadsheet surface, then lets an admin turn the important parts of a calculator into a controlled workflow:

1. Build a spreadsheet calculator in Quoin or import an existing `.xlsx` workbook.
2. Verify the math in the grid before adding workflow behavior.
3. Name the cells that matter, promoting them into Smart Cells.
4. Configure each Smart Cell's role, type, labels, rules, lookup behavior, and runner-facing behavior.
5. Use Runner Preview to generate the controlled form from the surfaced Smart Cells.

Existing `.xlsx` calculators import directly, so admins can start from working shop calculators instead of rebuilding them from scratch.

## Current State

Quoin is currently an internal prototype under active development. It is browser-local, with no backend and no user accounts. The beam selection demo shows the intended shape of the workflow: spreadsheet authoring first, then structured Smart Cells, then a runner-facing controlled form.

## Stack

Next.js · TypeScript · React · browser local storage

## Running It

On Windows, double-click:

```bat
Quoin.bat
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

## Testing

Available test commands:

```bash
npm run typecheck
npm run test:engine
npm run test:import
npm run test:import-reader
npm run fixtures:smoke
```
