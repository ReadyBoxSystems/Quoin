import fs from "node:fs/promises";
import { convertImportedSheetToQuoin } from "../lib/import/convert.ts";
import { readExcelWorkbook } from "../lib/import/read.ts";

const fixtureFiles = [
  "import-test-01-basic-values.xlsx",
  "import-test-02-basic-formulas.xlsx",
  "import-test-03-named-cells.xlsx",
  "import-test-04-multi-sheet-cross-sheet.xlsx",
  "import-test-05-reference-table-style.xlsx",
  "import-test-06-review-warnings.xlsx",
  "import-test-07-workbook-sheets.xlsx",
  "import-test-08-open-selected-sheet.xlsx",
  "import-test-09-names-across-sheets.xlsx",
  "import-test-10-sheet-names-with-spaces.xlsx",
  "import-test-11-sparse-large-sheet.xlsx",
  "import-test-12-workbook-review-mix.xlsx",
];

for (const fileName of fixtureFiles) {
  const buffer = await fs.readFile(fileName);
  const data = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
  const workbook = await readExcelWorkbook(fileName, data);

  console.log(`${fileName}: ${workbook.sheets.length} sheet(s)`);
  for (const sheet of workbook.sheets) {
    const converted = convertImportedSheetToQuoin(sheet, { names: workbook.names });
    const formulaCount = sheet.cells.filter((cell) => cell.kind === "formula").length;
    const reviewCount = workbook.reviewItems.length + converted.reviewItems.length;

    console.log(
      `  ${sheet.name}: ${Object.keys(converted.cells).length} cells, ${formulaCount} formula(s), ${converted.promotedNameCount} promoted name(s), ${reviewCount} review item(s)`,
    );
  }
}
