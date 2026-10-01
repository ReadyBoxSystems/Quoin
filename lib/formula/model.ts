export type FormulaReferenceKind = "cell" | "range" | "name";

export interface FormulaReference {
  kind: FormulaReferenceKind;
  raw: string;
  start: number;
  end: number;
  sheetName?: string;
  address?: string;
  rangeEnd?: string;
}

export interface ParsedFormula {
  source: string;
  references: FormulaReference[];
  functions: string[];
}

const cellPattern = /^\$?[A-Z]+\$?[1-9]\d*$/i;
const namePattern = /^[A-Za-z_][A-Za-z0-9_]*$/;
const formulaKeywords = new Set(["true", "false", "null", "and", "or", "not"]);

export function parseFormula(source: string): ParsedFormula {
  const references: FormulaReference[] = [];
  const functions: string[] = [];
  let index = source.startsWith("=") ? 1 : 0;

  while (index < source.length) {
    const char = source[index];
    if (char === '"') {
      index++;
      while (index < source.length) {
        if (source[index] === '"' && source[index + 1] === '"') index += 2;
        else if (source[index] === '"') { index++; break; }
        else index++;
      }
      continue;
    }

    const start = index;
    let sheetName: string | undefined;
    if (char === "'") {
      index++;
      let value = "";
      while (index < source.length) {
        if (source[index] === "'" && source[index + 1] === "'") { value += "'"; index += 2; }
        else if (source[index] === "'") { index++; break; }
        else value += source[index++];
      }
      if (source[index] === "!") { sheetName = value; index++; }
      else continue;
    }

    const tokenStart = index;
    while (index < source.length && /[A-Za-z0-9_$.]/.test(source[index])) index++;
    if (tokenStart === index) { index++; continue; }
    const token = source.slice(tokenStart, index);
    if (!sheetName && source[index] === "!" && namePattern.test(token)) {
      sheetName = token;
      index++;
      const addressStart = index;
      while (index < source.length && /[A-Za-z0-9_$]/.test(source[index])) index++;
      const address = source.slice(addressStart, index);
      if (!cellPattern.test(address)) continue;
      readCellOrRange(source, start, index, sheetName, address, references, (next) => { index = next; });
      continue;
    }
    if (cellPattern.test(token)) {
      readCellOrRange(source, start, index, sheetName, token, references, (next) => { index = next; });
      continue;
    }
    if (source.slice(index).match(/^\s*\(/)) {
      functions.push(token.toUpperCase());
      continue;
    }
    if (namePattern.test(token) && !formulaKeywords.has(token.toLowerCase())) {
      references.push({ kind: "name", raw: source.slice(start, index), start, end: index });
    }
  }
  return { source, references, functions: [...new Set(functions)] };
}

function readCellOrRange(
  source: string,
  start: number,
  current: number,
  sheetName: string | undefined,
  address: string,
  references: FormulaReference[],
  setIndex: (next: number) => void,
) {
  let index = current;
  const colon = source.slice(index).match(/^\s*:\s*/);
  if (colon) {
    index += colon[0].length;
    const rangeStart = index;
    while (index < source.length && /[A-Za-z0-9_$]/.test(source[index])) index++;
    const rangeEnd = source.slice(rangeStart, index);
    if (cellPattern.test(rangeEnd)) {
      references.push({ kind: "range", raw: source.slice(start, index), start, end: index, sheetName, address, rangeEnd });
      setIndex(index);
      return;
    }
  }
  references.push({ kind: "cell", raw: source.slice(start, current), start, end: current, sheetName, address });
  setIndex(current);
}

export function rewriteFormula(
  source: string,
  rewrite: (reference: FormulaReference) => string | null,
): string {
  const parsed = parseFormula(source);
  let next = source;
  for (const reference of [...parsed.references].reverse()) {
    const replacement = rewrite(reference);
    if (replacement === null || replacement === reference.raw) continue;
    next = `${next.slice(0, reference.start)}${replacement}${next.slice(reference.end)}`;
  }
  return next;
}

export function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}
