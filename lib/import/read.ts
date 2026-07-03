import ExcelJS from "exceljs/dist/exceljs.min.js";
import type { Cell, CellValue, Worksheet } from "exceljs";
import type { ImportedCell, ImportedCellValue, ImportedDataValidation, ImportedMerge, ImportedName, ImportedNameKind, ImportedWorkbook, ImportReviewItem } from "./types";

interface DefinedNameRange {
  name: string;
  ranges: string[];
}

interface WorkbookReadContext {
  names: ImportedName[];
  worksheets: Worksheet[];
  reviewItems: ImportReviewItem[];
}

export async function readExcelWorkbook(fileName: string, data: ArrayBuffer): Promise<ImportedWorkbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(data);

  const reviewItems: ImportReviewItem[] = [];
  const names = readDefinedNames(workbook.definedNames.model);
  const context = { names, worksheets: workbook.worksheets, reviewItems };
  const sheets = workbook.worksheets.map((worksheet) => readWorksheet(worksheet, context));

  return {
    fileName,
    sheets,
    names,
    reviewItems,
  };
}

function readWorksheet(worksheet: Worksheet, context: WorkbookReadContext) {
  const cells: ImportedCell[] = [];
  const merges = readMergedRanges(worksheet);
  const dataValidations = readDataValidations(worksheet, context);
  const mergedCoveredCells = new Set(merges.flatMap((merge) => addressesInRange(merge.range).filter((address) => address !== merge.topLeft)));
  let maxRow = 0;
  let maxColumn = 0;

  for (const merge of merges) {
    const bottomRight = parseCellAddress(merge.bottomRight);
    if (bottomRight) {
      maxRow = Math.max(maxRow, bottomRight.row);
      maxColumn = Math.max(maxColumn, bottomRight.column);
    }
    context.reviewItems.push({
      severity: "info",
      sheetName: worksheet.name,
      address: merge.topLeft,
      message: `Merged Excel range ${merge.range} was detected. Quoin imported the top-left cell only; covered cells were left blank for review.`,
    });
  }

  worksheet.eachRow((row, rowNumber) => {
    maxRow = Math.max(maxRow, rowNumber);
    row.eachCell((cell, columnNumber) => {
      maxColumn = Math.max(maxColumn, columnNumber);
      if (mergedCoveredCells.has(cell.address.toUpperCase())) return;
      const importedCell = readCell(cell, worksheet.name, context.reviewItems);
      if (importedCell.kind !== "blank") cells.push(importedCell);
    });
  });

  for (const validation of dataValidations) {
    const parsed = parseCellAddress(validation.address);
    if (!parsed) continue;
    maxRow = Math.max(maxRow, parsed.row);
    maxColumn = Math.max(maxColumn, parsed.column);
  }

  return {
    id: String(worksheet.id),
    name: worksheet.name,
    dimensions: {
      rowCount: maxRow,
      columnCount: maxColumn,
    },
    cells,
    merges,
    dataValidations,
  };
}

function readDataValidations(worksheet: Worksheet, context: WorkbookReadContext): ImportedDataValidation[] {
  const worksheetWithValidations = worksheet as Worksheet & {
    dataValidations?: { model?: Record<string, { type?: string; formulae?: unknown[] }> };
  };
  const model = worksheetWithValidations.dataValidations?.model;
  if (!model || typeof model !== "object") return [];

  const validations: ImportedDataValidation[] = [];

  for (const [target, validation] of Object.entries(model)) {
    if (validation?.type !== "list") continue;
    const formula = validation.formulae?.[0];
    const source = typeof formula === "string" ? formula.trim() : formula === undefined || formula === null ? "" : String(formula);
    const options = optionsFromLiteralListFormula(source) ?? optionsFromWorkbookSource(source, worksheet.name, context);

    for (const address of addressesForValidationTarget(target)) {
      validations.push({
        address,
        type: "list",
        ...(options ? { options } : { source }),
      });
    }
  }

  return validations;
}

function addressesForValidationTarget(target: string): string[] {
  const normalizedTarget = target.replace(/\$/g, "").toUpperCase();
  if (normalizedTarget.includes(":")) return addressesInRange(normalizedTarget);
  const address = normalizeAddress(normalizedTarget);
  return address ? [address] : [];
}

function optionsFromLiteralListFormula(formula: string): string[] | null {
  const trimmed = formula.trim().replace(/^=/, "").trim();
  if (!trimmed.startsWith('"') || !trimmed.endsWith('"')) return null;

  const literal = trimmed.slice(1, -1).replace(/""/g, '"');
  const options = literal
    .split(",")
    .map((option) => option.trim())
    .filter(Boolean);

  return options.length > 0 ? options : null;
}

function optionsFromWorkbookSource(source: string, currentSheetName: string, context: WorkbookReadContext): string[] | null {
  const sourceRange = rangeFromValidationSource(source, currentSheetName, context.names);
  if (!sourceRange) return null;

  const sourceWorksheet = context.worksheets.find((worksheet) => worksheet.name === sourceRange.sheetName);
  if (!sourceWorksheet) return null;

  const options: string[] = [];
  const seen = new Set<string>();

  for (const address of addressesInRange(sourceRange.range)) {
    const cell = sourceWorksheet.getCell(address);
    const value = dropdownOptionValueFromCellValue(cell.value, sourceWorksheet.name, address, context.reviewItems);
    if (value === null) continue;
    if (seen.has(value)) continue;
    seen.add(value);
    options.push(value);
  }

  return options.length > 0 ? options : null;
}

function rangeFromValidationSource(
  source: string,
  currentSheetName: string,
  names: ImportedName[],
): { sheetName: string; range: string } | null {
  const trimmed = source.trim().replace(/^=/, "").trim();
  if (!trimmed) return null;
  if (trimmed.includes("[") || trimmed.includes("]")) return null;
  if (trimmed.includes(",") || trimmed.includes(";")) return null;

  const directRange = sheetRangeFromReference(trimmed, currentSheetName);
  if (directRange) return directRange;

  const nameMatches = names.filter((name) => name.name.toLowerCase() === trimmed.toLowerCase());
  if (nameMatches.length !== 1 || nameMatches[0].kind !== "range") return null;

  return sheetRangeFromReference(nameMatches[0].reference, nameMatches[0].sheetName ?? currentSheetName);
}

function sheetRangeFromReference(reference: string, currentSheetName: string): { sheetName: string; range: string } | null {
  const bangIndex = reference.lastIndexOf("!");
  const sheetName = bangIndex === -1 ? currentSheetName : unquoteSheetName(reference.slice(0, bangIndex));
  const rawRange = bangIndex === -1 ? reference : reference.slice(bangIndex + 1);
  const parsed = parseRange(rawRange);
  return parsed ? { sheetName, range: parsed.range } : null;
}

function unquoteSheetName(sheetName: string): string {
  const trimmed = sheetName.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

function dropdownOptionValueFromCellValue(
  value: CellValue,
  sheetName: string,
  address: string,
  reviewItems: ImportReviewItem[],
): string | null {
  const formulaResult = importedValueFromFormulaResult(value, sheetName, address, reviewItems);
  const importedValue = formulaResult ?? importedValueFromCellValue(value, sheetName, address, reviewItems);
  if (importedValue === undefined || importedValue === null) return null;
  if (importedValue instanceof Date) return importedValue.toISOString().slice(0, 10);
  const option = String(importedValue);
  return option.trim() === "" ? null : option;
}

function readMergedRanges(worksheet: Worksheet): ImportedMerge[] {
  const model = worksheet.model as { merges?: string[] };
  const ranges = Array.isArray(model.merges) ? model.merges : [];
  return ranges
    .map((range) => {
      const parsed = parseRange(range);
      return parsed ? { range: parsed.range, topLeft: parsed.topLeft, bottomRight: parsed.bottomRight } : null;
    })
    .filter((merge): merge is ImportedMerge => merge !== null);
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

function parseRange(range: string): { range: string; topLeft: string; bottomRight: string } | null {
  const [start, end] = range.split(":").map((part) => normalizeAddress(part));
  if (!start || !end) return null;
  return {
    range: `${start}:${end}`,
    topLeft: start,
    bottomRight: end,
  };
}

function addressesInRange(range: string): string[] {
  const parsed = parseRange(range);
  if (!parsed) return [];

  const start = parseCellAddress(parsed.topLeft);
  const end = parseCellAddress(parsed.bottomRight);
  if (!start || !end) return [];

  const firstColumn = Math.min(start.column, end.column);
  const lastColumn = Math.max(start.column, end.column);
  const firstRow = Math.min(start.row, end.row);
  const lastRow = Math.max(start.row, end.row);
  const addresses: string[] = [];

  for (let row = firstRow; row <= lastRow; row += 1) {
    for (let column = firstColumn; column <= lastColumn; column += 1) {
      addresses.push(`${columnName(column)}${row}`);
    }
  }

  return addresses;
}

function normalizeAddress(address: string): string | null {
  const normalized = address.replace(/\$/g, "").toUpperCase();
  return /^[A-Z]+[1-9]\d*$/.test(normalized) ? normalized : null;
}

function parseCellAddress(address: string): { column: number; row: number } | null {
  const normalized = normalizeAddress(address);
  const match = normalized?.match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) return null;
  return {
    column: columnNumber(match[1]),
    row: Number(match[2]),
  };
}

function columnNumber(column: string): number {
  return column.split("").reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function columnName(column: number): string {
  let remaining = column;
  let name = "";

  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    remaining = Math.floor((remaining - modulo) / 26);
  }

  return name;
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
