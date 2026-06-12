import assert from "node:assert/strict";
import { convertImportedSheetToQuoin } from "./convert";
import type { ImportedSheet } from "./types";

function runTest(name: string, test: () => void) {
  try {
    test();
    console.log(`  ok - ${name}`);
  } catch (error) {
    console.error(`  fail - ${name}`);
    throw error;
  }
}

function makeSheet(partial: Partial<ImportedSheet>): ImportedSheet {
  return {
    id: "sheet-1",
    name: "Calculator",
    dimensions: { rowCount: 1, columnCount: 1 },
    cells: [],
    ...partial,
  };
}

console.log("convertImportedSheetToQuoin");

runTest("imports values and formulas as Quoin grid cell entries", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      cells: [
        { address: "A1", kind: "value", value: "Span" },
        { address: "B1", kind: "value", value: 14 },
        { address: "C1", kind: "formula", formula: "B1 * 2", value: 28 },
        { address: "D1", kind: "formula", formula: "=SUM(B1:C1)", value: 42 },
      ],
    }),
  );

  assert.equal(converted.cells.A1.entry, "Span");
  assert.equal(converted.cells.A1.type, "text");
  assert.equal(converted.cells.B1.entry, "14");
  assert.equal(converted.cells.B1.type, "number");
  assert.equal(converted.cells.C1.entry, "=B1 * 2");
  assert.equal(converted.cells.C1.role, "formula");
  assert.equal(converted.cells.D1.entry, "=SUM(B1:C1)");
});

runTest("expands dimensions to fit imported content", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      dimensions: { rowCount: 2, columnCount: 2 },
      cells: [{ address: "AA20", kind: "value", value: true }],
    }),
  );

  assert.equal(converted.columnCount, 27);
  assert.equal(converted.rowCount, 20);
  assert.equal(converted.cells.AA20.type, "boolean");
  assert.equal(converted.cells.AA20.entry, "true");
});

runTest("skips invalid cell addresses and adds a review item", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      cells: [{ address: "not-a-cell", kind: "value", value: "bad" }],
    }),
  );

  assert.deepEqual(Object.keys(converted.cells), []);
  assert.equal(converted.reviewItems.length, 1);
  assert.equal(converted.reviewItems[0].severity, "warning");
});

runTest("promotes safe single-cell workbook names to Smart Cell names", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      cells: [{ address: "B2", kind: "value", value: 12 }],
    }),
    {
      names: [{ name: "design_span", kind: "singleCell", reference: "'Calculator'!$B$2", sheetName: "Calculator" }],
    },
  );

  assert.equal(converted.promotedNameCount, 1);
  assert.equal(converted.cells.B2.name, "design_span");
  assert.equal(converted.cells.B2.label, "Design Span");
  assert.equal(converted.cells.B2.surfaced, false);
});

runTest("does not promote unsafe workbook names", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      cells: [{ address: "B2", kind: "value", value: 12 }],
    }),
    {
      names: [{ name: "Design Span", kind: "singleCell", reference: "B2", sheetName: "Calculator" }],
    },
  );

  assert.equal(converted.promotedNameCount, 0);
  assert.equal(converted.cells.B2.name, "");
  assert.equal(converted.reviewItems.length, 1);
  assert.match(converted.reviewItems[0].message, /not a safe Smart Cell name/);
});

runTest("reports ranges without promoting them", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      cells: [{ address: "A1", kind: "value", value: "table" }],
    }),
    {
      names: [{ name: "beam_table", kind: "range", reference: "A1:C10", sheetName: "Calculator" }],
    },
  );

  assert.equal(converted.promotedNameCount, 0);
  assert.equal(converted.reviewItems.length, 1);
  assert.equal(converted.reviewItems[0].severity, "info");
});

runTest("does not promote names that reference a different sheet", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      cells: [{ address: "B2", kind: "value", value: 12 }],
    }),
    {
      names: [{ name: "other_span", kind: "singleCell", reference: "'Other Sheet'!$B$2" }],
    },
  );

  assert.equal(converted.promotedNameCount, 0);
  assert.equal(converted.cells.B2.name, "");
  assert.equal(converted.reviewItems.length, 1);
});

runTest("flags imported formulas that need compatibility review", () => {
  const converted = convertImportedSheetToQuoin(
    makeSheet({
      cells: [
        { address: "A1", kind: "formula", formula: "Inputs!B2*2" },
        { address: "A2", kind: "formula", formula: "SUM(Table1[Amount])" },
        { address: "A3", kind: "formula", formula: "SUM('[Other.xlsx]Sheet1'!A1)" },
        { address: "A4", kind: "formula", formula: "SUM(A1;A2)" },
        { address: "A5", kind: "formula", formula: "A1#" },
      ],
    }),
  );

  assert.equal(converted.cells.A1.entry, "=Inputs!B2*2");
  assert.equal(converted.reviewItems.filter((item) => item.message.includes("cross-sheet reference")).length, 2);
  assert.equal(converted.reviewItems.filter((item) => item.message.includes("structured table reference")).length, 1);
  assert.equal(converted.reviewItems.filter((item) => item.message.includes("external workbook reference")).length, 1);
  assert.equal(converted.reviewItems.filter((item) => item.message.includes("semicolon argument separator")).length, 1);
  assert.equal(converted.reviewItems.filter((item) => item.message.includes("dynamic array")).length, 1);
});

console.log("All import converter tests passed");
