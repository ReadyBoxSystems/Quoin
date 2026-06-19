import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { readExcelWorkbook } from "./read";

async function runTest(name: string, test: () => Promise<void> | void) {
  try {
    await test();
    console.log(`  ok - ${name}`);
  } catch (error) {
    console.error(`  fail - ${name}`);
    throw error;
  }
}

async function workbookToArrayBuffer(workbook: ExcelJS.Workbook): Promise<ArrayBuffer> {
  const buffer = await workbook.xlsx.writeBuffer();
  const view = new Uint8Array(buffer);
  return view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength);
}

console.log("readExcelWorkbook");

async function main() {
  await runTest("reads sheets, cells, formulas, rich text, and defined names", async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Calculator");

    worksheet.getCell("A1").value = "Span";
    worksheet.getCell("B1").value = 14;
    worksheet.getCell("C1").value = { formula: "B1*2", result: 28 };
    worksheet.getCell("D1").value = {
      richText: [
        { text: "Shop" },
        { text: " note" },
      ],
    };

    workbook.definedNames.add("'Calculator'!$B$1", "design_span");
    workbook.definedNames.add("'Calculator'!$A$1:$D$1", "header_range");

    const imported = await readExcelWorkbook("calculator.xlsx", await workbookToArrayBuffer(workbook));
    const sheet = imported.sheets[0];

    assert.equal(imported.fileName, "calculator.xlsx");
    assert.equal(imported.sheets.length, 1);
    assert.equal(sheet.name, "Calculator");
    assert.equal(sheet.dimensions.rowCount, 1);
    assert.equal(sheet.dimensions.columnCount, 4);
    assert.deepEqual(
      sheet.cells.map((cell) => [cell.address, cell.kind, cell.value, cell.formula]),
      [
        ["A1", "value", "Span", undefined],
        ["B1", "value", 14, undefined],
        ["C1", "formula", 28, "B1*2"],
        ["D1", "value", "Shop note", undefined],
      ],
    );
    assert.deepEqual(
      imported.names.map((name) => [name.name, name.kind, name.reference, name.sheetName]),
      [
        ["design_span", "singleCell", "Calculator!$B$1", "Calculator"],
        ["header_range", "range", "Calculator!$A$1:$D$1", "Calculator"],
      ],
    );
  });

  await runTest("detects merged cells without importing covered duplicates", async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Merged");

    worksheet.getCell("A1").value = "Merged note";
    worksheet.mergeCells("A1:A3");
    worksheet.getCell("B1").value = "Normal";

    const imported = await readExcelWorkbook("merged.xlsx", await workbookToArrayBuffer(workbook));
    const sheet = imported.sheets[0];

    assert.deepEqual(sheet.merges, [{ range: "A1:A3", topLeft: "A1", bottomRight: "A3" }]);
    assert.deepEqual(
      sheet.cells.map((cell) => [cell.address, cell.kind, cell.value]),
      [
        ["A1", "value", "Merged note"],
        ["B1", "value", "Normal"],
      ],
    );
    assert.equal(sheet.dimensions.rowCount, 3);
    assert.equal(imported.reviewItems.some((item) => item.message.includes("Merged Excel range A1:A3")), true);
  });

  console.log("All workbook reader tests passed");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
