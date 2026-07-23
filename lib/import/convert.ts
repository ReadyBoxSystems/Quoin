import type { SmartCellType } from "@/lib/engine";
import type { GridCell } from "@/lib/sheet/types";
import type { ImportedCell, ImportedCellValue, ImportedDataValidation, ImportedName, ImportedSheet, ImportReviewItem } from "./types";

const defaultImportedColumnCount = 8;
const defaultImportedRowCount = 18;
const cellAddressPattern = /^([A-Z]+)([1-9]\d*)$/;
const safeSmartCellNamePattern = /^[A-Za-z_][A-Za-z0-9_]*$/;
const supportedFormulaFunctions = new Set([
  "SUM",
  "COUNT",
  "AVERAGE",
  "MAX",
  "MIN",
  "ROUND",
  "ROUNDUP",
  "ABS",
  "SQRT",
  "CEIL",
  "FLOOR",
  "IF",
  "VLOOKUP",
  "XLOOKUP",
]);

export interface ConvertedSheet {
  cells: Record<string, GridCell>;
  columnCount: number;
  rowCount: number;
  reviewItems: ImportReviewItem[];
  promotedNameCount: number;
}

export interface ConvertImportedSheetOptions {
  names?: ImportedName[];
  workbookSheets?: ImportedSheet[];
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
  const formulaCount = sheet.cells.filter((cell) => cell.kind === "formula").length;

  if (sheet.dimensions.rowCount >= 1000 && formulaCount === 0) {
    reviewItems.push({
      severity: "info",
      sheetName: sheet.name,
      message: `Sheet "${sheet.name}" looks like reference data (${sheet.dimensions.rowCount} rows, ${sheet.dimensions.columnCount} columns, no formulas). Quoin preserved it as a Sheet for now. Future Reference Table support should bind lookups to this data instead of surfacing it in Runner Preview.`,
    });
  }

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

  applyImportedDataValidations({
    cells,
    dataValidations: sheet.dataValidations ?? [],
    sheetName: sheet.name,
    reviewItems,
  });

  for (const validation of sheet.dataValidations ?? []) {
    const address = normalizeCellAddress(validation.address);
    if (!address) continue;
    const position = parseCellAddress(address);
    maxColumnCount = Math.max(maxColumnCount, position.columnNumber);
    maxRowCount = Math.max(maxRowCount, position.rowNumber);
  }

  return {
    cells,
    columnCount: maxColumnCount,
    rowCount: maxRowCount,
    reviewItems,
    promotedNameCount,
  };
}

function applyImportedDataValidations(input: {
  cells: Record<string, GridCell>;
  dataValidations: ImportedDataValidation[];
  sheetName: string;
  reviewItems: ImportReviewItem[];
}) {
  for (const validation of input.dataValidations) {
    const address = normalizeCellAddress(validation.address);
    if (!address) {
      input.reviewItems.push({
        severity: "warning",
        sheetName: input.sheetName,
        address: validation.address,
        message: `Skipped imported dropdown with unsupported address "${validation.address}".`,
      });
      continue;
    }

    if (!validation.options?.length) {
      input.reviewItems.push({
        severity: "info",
        sheetName: input.sheetName,
        address,
        message: `Excel dropdown at ${address} uses an unsupported source${validation.source ? ` (${validation.source})` : ""}. Quoin preserved the cell but needs future reference-data-backed dropdown support for this source.`,
      });
      continue;
    }

    const existingCell = input.cells[address] ?? makeEmptyGridCell(address);
    const generatedIdentity = existingCell.name
      ? {}
      : generatedDropdownIdentity(address, input.cells);

    input.cells[address] = {
      ...existingCell,
      role: existingCell.entry.trim().startsWith("=") ? existingCell.role : "input",
      type: existingCell.type === "number" || existingCell.type === "boolean" ? existingCell.type : "text",
      inputControl: "dropdown",
      inputOptions: validation.options,
      surfaced: true,
      ...generatedIdentity,
    };
  }
}

function generatedDropdownIdentity(address: string, cells: Record<string, GridCell>): Pick<GridCell, "name" | "label"> {
  const label = dropdownLabelForAddress(address, cells);
  const baseName = safeNameFromLabel(label) || `dropdown_${address.toLowerCase()}`;
  return {
    name: uniqueSmartCellName(baseName, cells),
    label,
  };
}

function dropdownLabelForAddress(address: string, cells: Record<string, GridCell>): string {
  const parsed = parseCellAddress(address);
  const leftColumn = columnNumberToLetters(parsed.columnNumber - 1);
  const leftAddress = leftColumn ? `${leftColumn}${parsed.rowNumber}` : "";
  const leftEntry = leftAddress ? cells[leftAddress]?.entry.trim() : "";
  return leftEntry || `Dropdown ${address}`;
}

function safeNameFromLabel(label: string): string {
  return label
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/^([0-9])/, "_$1");
}

function uniqueSmartCellName(baseName: string, cells: Record<string, GridCell>): string {
  const existingNames = new Set(Object.values(cells).map((cell) => cell.name).filter(Boolean));
  let candidate = safeSmartCellNamePattern.test(baseName) && !looksLikeCellAddress(baseName) ? baseName : `dropdown_${baseName}`;
  let suffix = 2;

  while (existingNames.has(candidate)) {
    candidate = `${baseName}_${suffix}`;
    suffix += 1;
  }

  return candidate;
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

  for (const vlookup of findFunctionCalls(formula, "VLOOKUP")) {
    const args = splitFormulaArguments(vlookup.args);
    const lookupValue = args[0] ?? "";
    const tableRange = args[1] ?? "";
    const outputColumn = args[2] ?? "";
    const matchMode = args[3] ?? "";
    const exactMatch = /^(false|0)$/i.test(matchMode.trim());
    const omittedMatchMode = args.length < 4 || matchMode.trim() === "";

    reviewItems.push({
      severity: exactMatch ? "info" : "warning",
      sheetName,
      address,
      formula: normalizeFormulaEntry(formula),
      message: exactMatch
        ? `Excel VLOOKUP at ${address} was preserved as a normal formula and should evaluate as an exact-match lookup. Source range: ${tableRange}; lookup key: ${lookupValue}; output column index: ${outputColumn}. Longer term, important lookups can be rebuilt as Quoin Reference Table or lookup Smart Cell logic for auditability, but promotion is not required for calculation.`
        : `Excel VLOOKUP at ${address} needs review before Quoin can evaluate it. Source range: ${tableRange || "unknown"}; lookup key: ${lookupValue || "unknown"}; output column index: ${outputColumn || "unknown"}. ${omittedMatchMode ? "Excel omits the match-mode argument here, which defaults to approximate matching." : `Match mode is ${matchMode}.`} Quoin currently supports exact-match VLOOKUP only; change the fourth argument to FALSE if exact matching is intended, or rebuild the behavior as an explicit Reference Table rule.`,
    });
  }

  for (const xlookup of findFunctionCalls(formula, "XLOOKUP")) {
    const args = splitFormulaArguments(xlookup.args);
    const lookupValue = args[0] ?? "";
    const lookupRange = args[1] ?? "";
    const returnRange = args[2] ?? "";
    const matchMode = args[4] ?? "";
    const exactMatch = matchMode.trim() === "" || matchMode.trim() === "0";

    reviewItems.push({
      severity: exactMatch ? "info" : "warning",
      sheetName,
      address,
      formula: normalizeFormulaEntry(formula),
      message: exactMatch
        ? `Excel XLOOKUP at ${address} was preserved as a normal formula and should evaluate as an exact-match lookup. Lookup key: ${lookupValue}; lookup range: ${lookupRange}; return range: ${returnRange}. Longer term, important lookups can be rebuilt as Quoin Reference Table or lookup Smart Cell logic for auditability.`
        : `Excel XLOOKUP at ${address} needs review before Quoin can evaluate it. Lookup key: ${lookupValue || "unknown"}; lookup range: ${lookupRange || "unknown"}; return range: ${returnRange || "unknown"}; match mode: ${matchMode}. Quoin currently supports exact-match XLOOKUP only; use match mode 0 if exact matching is intended.`,
    });
  }

  for (const functionName of functionNamesInFormula(formula)) {
    if (supportedFormulaFunctions.has(functionName)) continue;
    const guidance = unsupportedFunctionGuidance(functionName);
    reviewItems.push({
      severity: guidance.severity,
      sheetName,
      address,
      formula: normalizeFormulaEntry(formula),
      message: guidance.message(address, functionName),
    });
  }

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

function functionNamesInFormula(formula: string): string[] {
  const names = new Set<string>();
  const pattern = /\b([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(formula))) {
    names.add(match[1].toUpperCase());
  }

  return [...names];
}

function unsupportedFunctionGuidance(functionName: string): {
  severity: ImportReviewItem["severity"];
  message: (address: string, functionName: string) => string;
} {
  if (functionName === "IFERROR" || functionName === "IFNA") {
    return {
      severity: "warning",
      message: (address, name) => `Excel ${name} at ${address} is preserved but not evaluated by Quoin yet. Repair path: inspect the wrapped expression, decide what fallback value is safe for this workflow, and replace it with an explicit IF rule or a Quoin validation/review message so failures are visible instead of hidden.`,
    };
  }

  if (functionName === "INDIRECT" || functionName === "OFFSET") {
    return {
      severity: "warning",
      message: (address, name) => `Excel ${name} at ${address} is preserved but cannot be safely translated because it builds references dynamically. Repair path: replace the dynamic reference with direct cell references, or move the selectable data into a Reference Table and bind the selection as lookup criteria.`,
    };
  }

  if (functionName === "XLOOKUP" || functionName === "HLOOKUP" || functionName === "LOOKUP") {
    return {
      severity: "warning",
      message: (address, name) => `Excel ${name} at ${address} is preserved but not evaluated by Quoin yet. Repair path: rebuild it as an exact-match Reference Table lookup with explicit criteria and output column. Approximate or fallback behavior needs manual confirmation before translation.`,
    };
  }

  if (functionName === "INDEX" || functionName === "MATCH") {
    return {
      severity: "info",
      message: (address, name) => `Excel ${name} at ${address} is preserved for review. Repair path: if this is an INDEX/MATCH lookup, convert the source range into a Reference Table and bind the MATCH inputs as lookup criteria.`,
    };
  }

  if (functionName === "SUMIF" || functionName === "SUMIFS" || functionName === "COUNTIF" || functionName === "COUNTIFS" || functionName === "AVERAGEIF" || functionName === "AVERAGEIFS") {
    return {
      severity: "info",
      message: (address, name) => `Excel ${name} at ${address} is preserved for review. Repair path: convert the criteria ranges into a Reference Table or helper calculation, then define explicit criteria and the aggregate output Quoin should calculate.`,
    };
  }

  if (["TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY"].includes(functionName)) {
    return {
      severity: "warning",
      message: (address, name) => `Excel date function ${name} at ${address} is preserved but not evaluated by Quoin. Repair path: replace it with a fixed input, a numeric date/code field, or a future Quoin date rule after date semantics are defined.`,
    };
  }

  return {
    severity: "warning",
    message: (address, name) => `Excel function ${name} at ${address} is preserved but not currently supported by Quoin. Repair path: replace it with supported arithmetic, IF, range functions, or model the logic as a Smart Cell/reference-table step.`,
  };
}

function findFunctionCalls(formula: string, functionName: string): Array<{ args: string }> {
  const calls: Array<{ args: string }> = [];
  const pattern = new RegExp(`\\b${functionName}\\s*\\(`, "gi");
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(formula))) {
    const openIndex = match.index + match[0].lastIndexOf("(");
    const closeIndex = matchingParenIndex(formula, openIndex);
    if (closeIndex === null) continue;
    calls.push({ args: formula.slice(openIndex + 1, closeIndex) });
    pattern.lastIndex = closeIndex + 1;
  }

  return calls;
}

function matchingParenIndex(value: string, openIndex: number): number | null {
  let depth = 0;
  let inString = false;

  for (let index = openIndex; index < value.length; index += 1) {
    const char = value[index];
    if (char === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") depth += 1;
    if (char === ")") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }

  return null;
}

function splitFormulaArguments(args: string): string[] {
  const parts: string[] = [];
  let current = "";
  let depth = 0;
  let inString = false;

  for (const char of args) {
    if (char === '"') {
      inString = !inString;
      current += char;
      continue;
    }
    if (!inString && char === "(") depth += 1;
    if (!inString && char === ")") depth -= 1;
    if (!inString && depth === 0 && char === ",") {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }

  parts.push(current.trim());
  return parts;
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
    inputControl: "freeText",
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
    inputControl: "freeText",
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

function columnNumberToLetters(columnNumber: number): string | null {
  if (columnNumber < 1) return null;

  let remaining = columnNumber;
  let result = "";

  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    result = String.fromCharCode(65 + modulo) + result;
    remaining = Math.floor((remaining - modulo) / 26);
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
