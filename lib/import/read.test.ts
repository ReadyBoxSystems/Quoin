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

  await runTest("reads list data validations for dropdown imports", async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Dropdowns");

    worksheet.getCell("A1").value = "Load band";
    worksheet.getCell("B1").value = "standard";
    worksheet.getCell("B1").dataValidation = {
      type: "list",
      allowBlank: false,
      formulae: ['"standard,heavy"'],
    };
    worksheet.getCell("C1").dataValidation = {
      type: "list",
      formulae: ["Reference!$A$1:$A$3"],
    };

    const imported = await readExcelWorkbook("dropdowns.xlsx", await workbookToArrayBuffer(workbook));
    const sheet = imported.sheets[0];

    assert.deepEqual(sheet.dataValidations, [
      { address: "B1", type: "list", options: ["standard", "heavy"] },
      { address: "C1", type: "list", source: "Reference!$A$1:$A$3" },
    ]);
    assert.equal(sheet.dimensions.columnCount, 3);
  });

  await runTest("snapshots workbook-backed list validations into dropdown options", async () => {
    const workbook = new ExcelJS.Workbook();
    const calculator = workbook.addWorksheet("Calculator");
    const lists = workbook.addWorksheet("Lists");
    const referenceData = workbook.addWorksheet("Reference Data");

    calculator.getCell("A2").value = "small";
    calculator.getCell("A3").value = "";
    calculator.getCell("A4").value = "small";
    calculator.getCell("A5").value = true;

    lists.getCell("A1").value = "standard";
    lists.getCell("A2").value = "heavy";
    lists.getCell("A3").value = "heavy";
    lists.getCell("A4").value = 12;

    referenceData.getCell("B1").value = "top floor";
    referenceData.getCell("B2").value = "first floor";
    referenceData.getCell("B3").value = "dropped header";

    workbook.definedNames.add("'Lists'!$A$1:$A$4", "load_band_options");

    calculator.getCell("B1").dataValidation = { type: "list", formulae: ["=$A$2:$A$5"] };
    calculator.getCell("C1").dataValidation = { type: "list", formulae: ["=Lists!$A$1:$A$4"] };
    calculator.getCell("D1").dataValidation = { type: "list", formulae: ["='Reference Data'!$B$1:$B$3"] };
    calculator.getCell("E1").dataValidation = { type: "list", formulae: ["=load_band_options"] };
    calculator.getCell("F1").dataValidation = { type: "list", formulae: ["=Table1[Band]"] };

    const imported = await readExcelWorkbook("dropdown-ranges.xlsx", await workbookToArrayBuffer(workbook));
    const sheet = imported.sheets.find((item) => item.name === "Calculator");

    assert.ok(sheet);
    assert.deepEqual(sheet.dataValidations, [
      { address: "B1", type: "list", options: ["small", "true"] },
      { address: "C1", type: "list", options: ["standard", "heavy", "12"] },
      { address: "D1", type: "list", options: ["top floor", "first floor", "dropped header"] },
      { address: "E1", type: "list", options: ["standard", "heavy", "12"] },
      { address: "F1", type: "list", source: "=Table1[Band]" },
    ]);
  });

  console.log("All workbook reader tests passed");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
