import ExcelJS from "exceljs";

const fixtureTimestamp = new Date(Date.UTC(2026, 5, 1));

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
  {
    fileName: "import-test-07-workbook-sheets.xlsx",
    build(workbook) {
      const inputs = workbook.addWorksheet("Inputs");
      inputs.getCell("A1").value = "Workbook sheet preservation";
      inputs.getCell("A3").value = "Span";
      inputs.getCell("B3").value = 15;
      inputs.getCell("A4").value = "PLF";
      inputs.getCell("B4").value = 640;

      const reference = workbook.addWorksheet("Reference Data");
      reference.getCell("A1").value = "Reference data";
      reference.getCell("A3").value = "span";
      reference.getCell("B3").value = "beam";
      [
        [12, "2x12 SPF"],
        [14, "9.25 LVL"],
        [16, "11.875 LVL"],
      ].forEach((row, index) => {
        const rowNumber = index + 4;
        reference.getCell(`A${rowNumber}`).value = row[0];
        reference.getCell(`B${rowNumber}`).value = row[1];
      });

      const calc = workbook.addWorksheet("Calculator");
      calc.getCell("A1").value = "Calculator";
      calc.getCell("A3").value = "Total load";
      calc.getCell("B3").value = { formula: "Inputs!B3*Inputs!B4", result: 9600 };
      calc.getCell("A4").value = "Recommendation";
      calc.getCell("B4").value = { formula: 'IF(Inputs!B3>14,"11.875 LVL","9.25 LVL")', result: "11.875 LVL" };

      workbook.definedNames.add("Inputs!$B$3", "design_span");
      workbook.definedNames.add("Inputs!$B$4", "design_plf");
      workbook.definedNames.add("'Reference Data'!$A$3:$B$6", "beam_reference");
      workbook.definedNames.add("Calculator!$B$3", "total_load");
    },
  },
  {
    fileName: "import-test-08-open-selected-sheet.xlsx",
    build(workbook) {
      const notes = workbook.addWorksheet("Notes");
      notes.getCell("A1").value = "Import should keep this sheet";
      notes.getCell("A3").value = "Instruction";
      notes.getCell("B3").value = "Choose Calculator as the open Sheet during import.";

      const calc = workbook.addWorksheet("Calculator");
      calc.getCell("A1").value = "Open this sheet first";
      calc.getCell("A3").value = "Width";
      calc.getCell("B3").value = 8;
      calc.getCell("A4").value = "Height";
      calc.getCell("B4").value = 10;
      calc.getCell("A5").value = "Area";
      calc.getCell("B5").value = { formula: "B3*B4", result: 80 };

      const reference = workbook.addWorksheet("Lookup");
      reference.getCell("A1").value = "Supporting lookup sheet";
      reference.getCell("A3").value = "key";
      reference.getCell("B3").value = "value";
      reference.getCell("A4").value = "standard";
      reference.getCell("B4").value = "Use normal review path";

      workbook.definedNames.add("Calculator!$B$3", "part_width");
      workbook.definedNames.add("Calculator!$B$4", "part_height");
      workbook.definedNames.add("Calculator!$B$5", "part_area");
    },
  },
  {
    fileName: "import-test-09-names-across-sheets.xlsx",
    build(workbook) {
      const geometry = workbook.addWorksheet("Geometry");
      geometry.getCell("A1").value = "Geometry inputs";
      geometry.getCell("A3").value = "Span";
      geometry.getCell("B3").value = 18;
      geometry.getCell("A4").value = "Depth";
      geometry.getCell("B4").value = 11.875;

      const loads = workbook.addWorksheet("Loads");
      loads.getCell("A1").value = "Load inputs";
      loads.getCell("A3").value = "Live PLF";
      loads.getCell("B3").value = 420;
      loads.getCell("A4").value = "Dead PLF";
      loads.getCell("B4").value = 180;
      loads.getCell("A5").value = "Total PLF";
      loads.getCell("B5").value = { formula: "B3+B4", result: 600 };

      const summary = workbook.addWorksheet("Summary");
      summary.getCell("A1").value = "Summary";
      summary.getCell("A3").value = "Total load";
      summary.getCell("B3").value = { formula: "Geometry!B3*Loads!B5", result: 10800 };

      workbook.definedNames.add("Geometry!$B$3", "span_ft");
      workbook.definedNames.add("Geometry!$B$4", "member_depth");
      workbook.definedNames.add("Loads!$B$5", "total_plf");
      workbook.definedNames.add("Summary!$B$3", "summary_total_load");
    },
  },
  {
    fileName: "import-test-10-sheet-names-with-spaces.xlsx",
    build(workbook) {
      const input = workbook.addWorksheet("Input Data");
      input.getCell("A1").value = "Sheet names with spaces";
      input.getCell("A3").value = "Quantity";
      input.getCell("B3").value = 6;
      input.getCell("A4").value = "Unit weight";
      input.getCell("B4").value = 125;

      const calc = workbook.addWorksheet("Calc Sheet");
      calc.getCell("A1").value = "Quoted sheet references";
      calc.getCell("A3").value = "Total weight";
      calc.getCell("B3").value = { formula: "'Input Data'!B3*'Input Data'!B4", result: 750 };
      calc.getCell("A4").value = "Check";
      calc.getCell("B4").value = { formula: 'IF(\'Input Data\'!B3>5,"Review","OK")', result: "Review" };

      workbook.definedNames.add("'Input Data'!$B$3", "input_quantity");
      workbook.definedNames.add("'Calc Sheet'!$B$3", "total_weight");
    },
  },
  {
    fileName: "import-test-11-sparse-large-sheet.xlsx",
    build(workbook) {
      const sheet = workbook.addWorksheet("Sparse Large");
      sheet.getCell("A1").value = "Sparse imported layout";
      sheet.getCell("A3").value = "Near input";
      sheet.getCell("B3").value = 10;
      sheet.getCell("K29").value = "Far input";
      sheet.getCell("L29").value = 22;
      sheet.getCell("K30").value = "Far formula";
      sheet.getCell("M30").value = { formula: "B3+L29", result: 32 };

      workbook.definedNames.add("'Sparse Large'!$L$29", "far_input");
      workbook.definedNames.add("'Sparse Large'!$M$30", "far_total");
    },
  },
  {
    fileName: "import-test-12-workbook-review-mix.xlsx",
    build(workbook) {
      const clean = workbook.addWorksheet("Clean Inputs");
      clean.getCell("A1").value = "Clean imported sheet";
      clean.getCell("A3").value = "Input A";
      clean.getCell("B3").value = 5;
      clean.getCell("A4").value = "Input B";
      clean.getCell("B4").value = 7;

      const warnings = workbook.addWorksheet("Warning Formulas");
      warnings.getCell("A1").value = "Workbook-level review aggregation";
      warnings.getCell("A3").value = "Cross-sheet";
      warnings.getCell("B3").value = { formula: "'Clean Inputs'!B3+'Clean Inputs'!B4", result: 12 };
      warnings.getCell("A4").value = "Structured table";
      warnings.getCell("B4").value = { formula: "SUM(Table2[Cost])", result: 100 };
      warnings.getCell("A5").value = "Semicolon separator";
      warnings.getCell("B5").value = { formula: "SUM(B3;B4)", result: 112 };

      workbook.definedNames.add("'Clean Inputs'!$B$3", "clean_input_a");
      workbook.definedNames.add("'Warning Formulas'!$B$3", "warning_cross_sheet_total");
      workbook.definedNames.add("'Clean Inputs'!$A$3:$B$4", "clean_input_block");
    },
  },
];

for (const fixture of fixtures) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Quoin fixture generator";
  workbook.created = fixtureTimestamp;
  workbook.modified = fixtureTimestamp;
  fixture.build(workbook);
  await workbook.xlsx.writeFile(fixture.fileName);
  console.log(`Wrote ${fixture.fileName}`);
}
