import type { SmartCellType } from "@/lib/engine";
import type { GridCell } from "@/lib/sheet/types";
import type { ImportedCell, ImportedCellValue, ImportedName, ImportedSheet, ImportReviewItem } from "./types";

const defaultImportedColumnCount = 8;
const defaultImportedRowCount = 18;
const cellAddressPattern = /^([A-Z]+)([1-9]\d*)$/;
const safeSmartCellNamePattern = /^[A-Za-z_][A-Za-z0-9_]*$/;

export interface ConvertedSheet {
  cells: Record<string, GridCell>;
  columnCount: number;
  rowCount: number;
  reviewItems: ImportReviewItem[];
  promotedNameCount: number;
}

export interface ConvertImportedSheetOptions {
  names?: ImportedName[];
  minimumColumnCount?: number;
  minimumRowCount?: number;
}

export function convertImportedSheetToQuoin(sheet: ImportedSheet, options: ConvertImportedSheetOptions = {}): ConvertedSheet {
  const minimumColumnCount = options.minimumColumnCount ?? defaultImportedColumnCount;
  const minimumRowCount = options.minimumRowCount ?? defaultImportedRowCount;
  const reviewItems: ImportReviewItem[] = [];
  const cells: Record<string, GridCell> = {};
  let maxColumnCount = Math.max(minimumColumnCount, sheet.dimensions.columnCount);
  let maxRowCount = Math.max(minimumRowCount, sheet.dimensions.rowCount);

  for (const importedCell of sheet.cells) {
    const address = normalizeCellAddress(importedCell.address);
    if (!address) {
      reviewItems.push({
        severity: "warning",
        sheetName: sheet.name,
        address: importedCell.address,
        message: `Skipped imported cell with unsupported address "${importedCell.address}".`,
      });
      continue;
    }

    const position = parseCellAddress(address);
    maxColumnCount = Math.max(maxColumnCount, position.columnNumber);
    maxRowCount = Math.max(maxRowCount, position.rowNumber);

    if (importedCell.kind === "blank") continue;

    cells[address] = makeImportedGridCell(address, importedCell);
    reviewImportedFormula(importedCell, sheet.name, address, reviewItems);
  }

  const promotedNameCount = applyImportedNames({
    cells,
    names: options.names ?? [],
    sheetName: sheet.name,
    reviewItems,
  });

  return {
    cells,
    columnCount: maxColumnCount,
    rowCount: maxRowCount,
    reviewItems,
    promotedNameCount,
  };
}

function reviewImportedFormula(
  importedCell: ImportedCell,
  sheetName: string,
  address: string,
  reviewItems: ImportReviewItem[],
) {
  if (importedCell.kind !== "formula" || !importedCell.formula) return;

  const formula = importedCell.formula;
  const warningMessages: string[] = [];

  if (/\[[^\]]+\][^!]*!/.test(formula)) {
    warningMessages.push("external workbook reference");
  }

  if (/[A-Za-z_][A-Za-z0-9_]*\[[^\]]+\]/.test(formula)) {
    warningMessages.push("structured table reference");
  }

  if (formula.includes("#") || formula.includes("@")) {
    warningMessages.push("dynamic array or implicit-intersection marker");
  }

  if (formula.includes(";")) {
    warningMessages.push("semicolon argument separator");
  }

  for (const message of warningMessages) {
    reviewItems.push({
      severity: "warning",
      sheetName,
      address,
      formula: normalizeFormulaEntry(formula),
      message: `Formula may need review: ${message}.`,
    });
  }
}

function makeImportedGridCell(address: string, importedCell: ImportedCell): GridCell {
  const entry = importedCell.kind === "formula" ? normalizeFormulaEntry(importedCell.formula ?? "") : stringifyImportedValue(importedCell.value);

  return {
    address,
    entry,
    name: "",
    label: "",
    role: entry.startsWith("=") ? "formula" : "input",
    type: inferSmartCellType(importedCell),
    inputOptions: [],
    surfaced: false,
    annotation: "",
    ruleMessage: "",
  };
}

function applyImportedNames(input: {
  cells: Record<string, GridCell>;
  names: ImportedName[];
  sheetName: string;
  reviewItems: ImportReviewItem[];
}): number {
  let promotedNameCount = 0;
  const usedNames = new Set<string>();

  for (const importedName of input.names) {
    if (importedName.sheetName && importedName.sheetName !== input.sheetName) continue;

    if (importedName.kind !== "singleCell") {
      input.reviewItems.push({
        severity: "info",
        sheetName: importedName.sheetName,
        name: importedName.name,
        message: `Workbook name "${importedName.name}" points to ${importedName.kind}; it was not promoted to a Smart Cell.`,
      });
      continue;
    }

    const address = addressFromReference(importedName.reference, input.sheetName);
    if (!address) {
      input.reviewItems.push({
        severity: "warning",
        sheetName: importedName.sheetName ?? input.sheetName,
        name: importedName.name,
        message: `Workbook name "${importedName.name}" has an unsupported reference "${importedName.reference}".`,
      });
      continue;
    }

    if (!safeSmartCellNamePattern.test(importedName.name) || looksLikeCellAddress(importedName.name)) {
      input.reviewItems.push({
        severity: "warning",
        sheetName: importedName.sheetName ?? input.sheetName,
        address,
        name: importedName.name,
        message: `Workbook name "${importedName.name}" is not a safe Smart Cell name and was not promoted.`,
      });
      continue;
    }

    if (usedNames.has(importedName.name)) {
      input.reviewItems.push({
        severity: "warning",
        sheetName: importedName.sheetName ?? input.sheetName,
        address,
        name: importedName.name,
        message: `Workbook name "${importedName.name}" was duplicated and only the first matching cell was promoted.`,
      });
      continue;
    }

    const existingCell = input.cells[address] ?? makeEmptyGridCell(address);
    input.cells[address] = {
      ...existingCell,
      name: importedName.name,
      label: prettifyName(importedName.name),
      role: existingCell.entry.startsWith("=") ? "formula" : "input",
    };
    usedNames.add(importedName.name);
    promotedNameCount += 1;
  }

  return promotedNameCount;
}

function normalizeFormulaEntry(formula: string): string {
  const trimmed = formula.trim();
  if (!trimmed) return "=";
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

function stringifyImportedValue(value: ImportedCellValue | undefined): string {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value);
}

function inferSmartCellType(importedCell: ImportedCell): SmartCellType {
  if (typeof importedCell.value === "boolean") return "boolean";
  if (typeof importedCell.value === "number") return "number";
  if (importedCell.kind === "formula") return "number";
  return "text";
}

function makeEmptyGridCell(address: string): GridCell {
  return {
    address,
    entry: "",
    name: "",
    label: "",
    role: "input",
    type: "text",
    inputOptions: [],
    surfaced: false,
    annotation: "",
    ruleMessage: "",
  };
}

function normalizeCellAddress(address: string): string | null {
  const normalized = address.replace(/\$/g, "").toUpperCase();
  return cellAddressPattern.test(normalized) ? normalized : null;
}

function parseCellAddress(address: string): { columnLetters: string; columnNumber: number; rowNumber: number } {
  const match = address.match(cellAddressPattern);
  if (!match) throw new Error(`Invalid cell address: ${address}`);

  return {
    columnLetters: match[1],
    columnNumber: columnLettersToNumber(match[1]),
    rowNumber: Number(match[2]),
  };
}

function columnLettersToNumber(columnLetters: string): number {
  let result = 0;
  for (const letter of columnLetters) {
    result = result * 26 + letter.charCodeAt(0) - 64;
  }
  return result;
}

function addressFromReference(reference: string, selectedSheetName: string): string | null {
  const trimmed = reference.trim();
  const parts = trimmed.split("!");
  const rawAddress = parts.length === 1 ? parts[0] : parts[parts.length - 1];
  const sheetName = parts.length === 1 ? selectedSheetName : unquoteSheetName(parts.slice(0, -1).join("!"));

  if (sheetName !== selectedSheetName) return null;
  return normalizeCellAddress(rawAddress);
}

function unquoteSheetName(sheetName: string): string {
  const trimmed = sheetName.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

function looksLikeCellAddress(value: string): boolean {
  return cellAddressPattern.test(value.toUpperCase());
}

function prettifyName(name: string): string {
  return name
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}
