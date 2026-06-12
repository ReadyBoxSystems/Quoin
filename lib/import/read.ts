import ExcelJS from "exceljs/dist/exceljs.min.js";
import type { Cell, CellValue, Worksheet } from "exceljs";
import type { ImportedCell, ImportedCellValue, ImportedName, ImportedNameKind, ImportedWorkbook, ImportReviewItem } from "./types";

interface DefinedNameRange {
  name: string;
  ranges: string[];
}

export async function readExcelWorkbook(fileName: string, data: ArrayBuffer): Promise<ImportedWorkbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(data);

  const reviewItems: ImportReviewItem[] = [];
  const names = readDefinedNames(workbook.definedNames.model);
  const sheets = workbook.worksheets.map((worksheet) => readWorksheet(worksheet, reviewItems));

  return {
    fileName,
    sheets,
    names,
    reviewItems,
  };
}

function readWorksheet(worksheet: Worksheet, reviewItems: ImportReviewItem[]) {
  const cells: ImportedCell[] = [];
  let maxRow = 0;
  let maxColumn = 0;

  worksheet.eachRow((row, rowNumber) => {
    maxRow = Math.max(maxRow, rowNumber);
    row.eachCell((cell, columnNumber) => {
      maxColumn = Math.max(maxColumn, columnNumber);
      const importedCell = readCell(cell, worksheet.name, reviewItems);
      if (importedCell.kind !== "blank") cells.push(importedCell);
    });
  });

  return {
    id: String(worksheet.id),
    name: worksheet.name,
    dimensions: {
      rowCount: maxRow,
      columnCount: maxColumn,
    },
    cells,
  };
}

function readCell(cell: Cell, sheetName: string, reviewItems: ImportReviewItem[]): ImportedCell {
  const value = cell.value;

  if (value === null || value === undefined) {
    return { address: cell.address, kind: "blank" };
  }

  const formula = formulaFromCellValue(value);
  if (formula) {
    return {
      address: cell.address,
      kind: "formula",
      formula,
      value: importedValueFromFormulaResult(value, sheetName, cell.address, reviewItems),
    };
  }

  const importedValue = importedValueFromCellValue(value, sheetName, cell.address, reviewItems);
  if (importedValue === undefined) {
    return { address: cell.address, kind: "blank" };
  }

  return {
    address: cell.address,
    kind: "value",
    value: importedValue,
  };
}

function formulaFromCellValue(value: CellValue): string | null {
  if (!value || typeof value !== "object") return null;
  if ("formula" in value && typeof value.formula === "string") return value.formula;
  if ("sharedFormula" in value && typeof value.formula === "string") return value.formula;
  return null;
}

function importedValueFromFormulaResult(
  value: CellValue,
  sheetName: string,
  address: string,
  reviewItems: ImportReviewItem[],
): ImportedCellValue | undefined {
  if (!value || typeof value !== "object" || !("result" in value)) return undefined;
  return importedValueFromCellValue(value.result as CellValue, sheetName, address, reviewItems);
}

function importedValueFromCellValue(
  value: CellValue,
  sheetName: string,
  address: string,
  reviewItems: ImportReviewItem[],
): ImportedCellValue | undefined {
  if (value === null || value === undefined) return undefined;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean" || value instanceof Date) return value;

  if (typeof value === "object") {
    if ("text" in value && typeof value.text === "string") return value.text;
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text ?? "").join("");
    }
    if ("error" in value && typeof value.error === "string") {
      reviewItems.push({
        severity: "warning",
        sheetName,
        address,
        message: `Imported Excel error value "${value.error}" as text.`,
      });
      return value.error;
    }
  }

  reviewItems.push({
    severity: "warning",
    sheetName,
    address,
    message: "Skipped a cell with an unsupported Excel value type.",
  });
  return undefined;
}

function readDefinedNames(model: DefinedNameRange[]): ImportedName[] {
  const names: ImportedName[] = [];

  for (const definition of model) {
    for (const reference of definition.ranges) {
      names.push({
        name: definition.name,
        reference,
        sheetName: sheetNameFromReference(reference),
        kind: classifyDefinedName(reference),
      });
    }
  }

  return names;
}

function classifyDefinedName(reference: string): ImportedNameKind {
  const trimmed = reference.trim();
  if (!trimmed) return "unknown";
  if (trimmed.includes("[") || trimmed.includes("]")) return "external";
  if (trimmed.includes(":")) return "range";
  if (looksLikeSheetCellReference(trimmed) || looksLikeCellReference(trimmed)) return "singleCell";
  if (trimmed.startsWith("=")) return "formula";
  return "unknown";
}

function sheetNameFromReference(reference: string): string | undefined {
  const bangIndex = reference.lastIndexOf("!");
  if (bangIndex === -1) return undefined;

  const rawSheetName = reference.slice(0, bangIndex).trim();
  if (rawSheetName.startsWith("'") && rawSheetName.endsWith("'")) {
    return rawSheetName.slice(1, -1).replace(/''/g, "'");
  }
  return rawSheetName || undefined;
}

function looksLikeSheetCellReference(reference: string): boolean {
  const bangIndex = reference.lastIndexOf("!");
  if (bangIndex === -1) return false;
  return looksLikeCellReference(reference.slice(bangIndex + 1));
}

function looksLikeCellReference(reference: string): boolean {
  return /^\$?[A-Z]+\$?[1-9]\d*$/i.test(reference.trim());
}
