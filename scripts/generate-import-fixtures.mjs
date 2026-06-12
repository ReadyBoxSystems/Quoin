import ExcelJS from "exceljs";

const fixtures = [
  {
    fileName: "import-test-01-basic-values.xlsx",
    build(workbook) {
      const sheet = workbook.addWorksheet("Basic Values");
      sheet.getCell("A1").value = "Basic value import";
      sheet.getCell("A3").value = "Label";
      sheet.getCell("B3").value = "Value";
      sheet.getCell("A4").value = "Design span";
      sheet.getCell("B4").value = 14;
      sheet.getCell("A5").value = "Load band";
      sheet.getCell("B5").value = "standard";
      sheet.getCell("A6").value = "Needs review";
      sheet.getCell("B6").value = true;
      sheet.getCell("A7").value = "Import date";
      sheet.getCell("B7").value = new Date(Date.UTC(2026, 5, 2));
    },
  },
  {
    fileName: "import-test-02-basic-formulas.xlsx",
    build(workbook) {
      const sheet = workbook.addWorksheet("Basic Formulas");
      sheet.getCell("A1").value = "Basic formula import";
      sheet.getCell("A3").value = "Span";
      sheet.getCell("B3").value = 14;
      sheet.getCell("A4").value = "PLF";
      sheet.getCell("B4").value = 650;
      sheet.getCell("A5").value = "Total load";
      sheet.getCell("B5").value = { formula: "B3*B4", result: 9100 };
      sheet.getCell("A6").value = "Rounded load";
      sheet.getCell("B6").value = { formula: "ROUND(B5,0)", result: 9100 };
      sheet.getCell("A7").value = "Pass check";
      sheet.getCell("B7").value = { formula: 'IF(B5<10000,"PASS","FAIL")', result: "PASS" };
    },
  },
  {
    fileName: "import-test-03-named-cells.xlsx",
    build(workbook) {
      const sheet = workbook.addWorksheet("Named Cells");
      sheet.getCell("A1").value = "Named cell import";
      sheet.getCell("A3").value = "Span";
      sheet.getCell("B3").value = 16;
      sheet.getCell("A4").value = "PLF";
      sheet.getCell("B4").value = 700;
      sheet.getCell("A5").value = "Total load";
      sheet.getCell("B5").value = { formula: "design_span*design_plf", result: 11200 };
      sheet.getCell("A7").value = "Unsafe name target";
      sheet.getCell("B7").value = "kept as normal cell";

      workbook.definedNames.add("'Named Cells'!$B$3", "design_span");
      workbook.definedNames.add("'Named Cells'!$B$4", "design_plf");
      workbook.definedNames.add("'Named Cells'!$A$3:$B$5", "input_block");
      workbook.definedNames.add("'Named Cells'!$B$7", "Design Span Unsafe");
    },
  },
  {
    fileName: "import-test-04-multi-sheet-cross-sheet.xlsx",
    build(workbook) {
      const inputs = workbook.addWorksheet("Inputs");
      inputs.getCell("A1").value = "Inputs";
      inputs.getCell("A3").value = "Span";
      inputs.getCell("B3").value = 12;
      inputs.getCell("A4").value = "PLF";
      inputs.getCell("B4").value = 625;

      const calc = workbook.addWorksheet("Calculator");
      calc.getCell("A1").value = "Calculator using another sheet";
      calc.getCell("A3").value = "Total load";
      calc.getCell("B3").value = { formula: "Inputs!B3*Inputs!B4", result: 7500 };
      calc.getCell("A4").value = "Review note";
      calc.getCell("B4").value = { formula: 'IF(Inputs!B3>14,"Review","OK")', result: "OK" };

      workbook.definedNames.add("Inputs!$B$3", "input_span");
      workbook.definedNames.add("Calculator!$B$3", "total_load");
    },
  },
  {
    fileName: "import-test-05-reference-table-style.xlsx",
    build(workbook) {
      const sheet = workbook.addWorksheet("Reference Table");
      sheet.getCell("A1").value = "Lookup table style import";
      sheet.getCell("A3").value = "span";
      sheet.getCell("B3").value = "load_band";
      sheet.getCell("C3").value = "beam";
      const rows = [
        [10, "standard", "2x10 SPF"],
        [12, "standard", "2x12 SPF"],
        [14, "standard", "9.25 LVL"],
        [16, "heavy", "14 LVL"],
      ];
      rows.forEach((row, index) => {
        const rowNumber = index + 4;
        sheet.getCell(`A${rowNumber}`).value = row[0];
        sheet.getCell(`B${rowNumber}`).value = row[1];
        sheet.getCell(`C${rowNumber}`).value = row[2];
      });
      sheet.getCell("E3").value = "Selected span";
      sheet.getCell("F3").value = 14;
      sheet.getCell("E4").value = "Formula result";
      sheet.getCell("F4").value = { formula: 'IF(F3=14,"9.25 LVL","manual review")', result: "9.25 LVL" };

      workbook.definedNames.add("'Reference Table'!$A$3:$C$7", "beam_table");
      workbook.definedNames.add("'Reference Table'!$F$3", "selected_span");
    },
  },
  {
    fileName: "import-test-06-review-warnings.xlsx",
    build(workbook) {
      const sheet = workbook.addWorksheet("Review Warnings");
      sheet.getCell("A1").value = "Formula review warning import";
      sheet.getCell("A3").value = "Cross-sheet formula";
      sheet.getCell("B3").value = { formula: "Inputs!B2*2", result: 20 };
      sheet.getCell("A4").value = "Structured table formula";
      sheet.getCell("B4").value = { formula: "SUM(Table1[Amount])", result: 100 };
      sheet.getCell("A5").value = "External workbook formula";
      sheet.getCell("B5").value = { formula: "SUM('[Other.xlsx]Sheet1'!A1)", result: 5 };
      sheet.getCell("A6").value = "Semicolon separator formula";
      sheet.getCell("B6").value = { formula: "SUM(A1;A2)", result: 0 };
      sheet.getCell("A7").value = "Spill marker formula";
      sheet.getCell("B7").value = { formula: "A1#", result: 0 };
    },
  },
];

for (const fixture of fixtures) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Quoin fixture generator";
  workbook.created = new Date();
  workbook.modified = new Date();
  fixture.build(workbook);
  await workbook.xlsx.writeFile(fixture.fileName);
  console.log(`Wrote ${fixture.fileName}`);
}
