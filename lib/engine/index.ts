import { all, create, type MathJsInstance, type MathNode } from "mathjs";
import type {
  CellValue,
  EngineCell,
  EngineInput,
  EngineIssue,
  EngineResult,
  RuleState,
  WorkbookEngineInput,
  WorkbookEngineResult,
  WorkbookEngineSheet,
  WorkbookEngineSheetResult,
} from "./types";

const math = create(all, {}) as MathJsInstance;

math.import(
  {
    SUM: (...values: unknown[]) => sumValues(values),
    average: (...values: unknown[]) => averageValues(values),
    AVERAGE: (...values: unknown[]) => averageValues(values),
    MAX: (...values: unknown[]) => Math.max(...numericValues(values)),
    MIN: (...values: unknown[]) => Math.min(...numericValues(values)),
    ROUND: (value: unknown, decimals?: unknown) => {
      const numericValue = toNumber(value) ?? 0;
      const places = toNumber(decimals) ?? 0;
      return Number(numericValue.toFixed(places));
    },
    ROUNDUP: (value: unknown, decimals?: unknown) => {
      const numericValue = toNumber(value) ?? 0;
      const places = Math.trunc(toNumber(decimals) ?? 0);
      const factor = 10 ** places;
      if (!Number.isFinite(factor) || factor === 0) return numericValue;
      return Math.sign(numericValue) * Math.ceil(Math.abs(numericValue) * factor) / factor;
    },
    roundup: (value: unknown, decimals?: unknown) => {
      const numericValue = toNumber(value) ?? 0;
      const places = Math.trunc(toNumber(decimals) ?? 0);
      const factor = 10 ** places;
      if (!Number.isFinite(factor) || factor === 0) return numericValue;
      return Math.sign(numericValue) * Math.ceil(Math.abs(numericValue) * factor) / factor;
    },
    ABS: (value: unknown) => Math.abs(toNumber(value) ?? 0),
    SQRT: (value: unknown) => Math.sqrt(toNumber(value) ?? 0),
    CEIL: (value: unknown) => Math.ceil(toNumber(value) ?? 0),
    FLOOR: (value: unknown) => Math.floor(toNumber(value) ?? 0),
    if: (condition: unknown, valueIfTrue: unknown, valueIfFalse: unknown) => (condition ? valueIfTrue : valueIfFalse),
    IF: (condition: unknown, valueIfTrue: unknown, valueIfFalse: unknown) => (condition ? valueIfTrue : valueIfFalse),
  },
  { override: true },
);

const ALLOWED_FUNCTIONS = new Set([
  "sum",
  "SUM",
  "mean",
  "average",
  "AVERAGE",
  "max",
  "MAX",
  "min",
  "MIN",
  "round",
  "ROUND",
  "ROUNDUP",
  "roundup",
  "abs",
  "ABS",
  "sqrt",
  "SQRT",
  "ceil",
  "CEIL",
  "floor",
  "FLOOR",
  "if",
  "IF",
]);

const DISALLOWED_NODE_TYPES = new Set([
  "AssignmentNode",
  "FunctionAssignmentNode",
  "AccessorNode",
  "ObjectNode",
  "ArrayNode",
  "IndexNode",
]);

export type {
  CellValue,
  EngineCell,
  EngineInput,
  EngineIssue,
  EngineResult,
  RuleState,
  SmartCellRole,
  SmartCellType,
  WorkbookEngineInput,
  WorkbookEngineResult,
  WorkbookEngineSheet,
  WorkbookEngineSheetResult,
} from "./types";

export function executeEngine(input: EngineInput): EngineResult {
  const cells = input.cells;
  const indexes = buildIndexes(cells);
  const errors: EngineIssue[] = [];
  const warnings: EngineIssue[] = [];
  const ruleStates: RuleState[] = [];
  const dependencies = new Map<string, Set<string>>();

  for (const cell of cells) {
    const refs = referencesForCell(cell);
    const deps = new Set<string>();

    for (const ref of refs) {
      const dependency = indexes.byReference.get(ref);
      if (!dependency) {
        if (isCellReference(ref) || isScopedCellReference(ref)) continue;
        errors.push(issue(cell, `Missing reference "${ref}".`));
        continue;
      }
      deps.add(dependency.id);
    }

    dependencies.set(cell.id, deps);
  }

  const sort = topologicalSort(cells, dependencies);
  for (const cycleId of sort.cycleIds) {
    const cell = indexes.byId.get(cycleId);
    if (cell) errors.push(issue(cell, "Cell participates in a circular dependency."));
  }

  if (errors.length > 0) {
    return emptyResult(false, cells, sort.order, errors, warnings, ruleStates);
  }

  const valuesById = new Map<string, CellValue>();
  const scope: Record<string, CellValue> = {};

  for (const cell of sort.order.map((id) => indexes.byId.get(id)).filter(Boolean) as EngineCell[]) {
    const value = evaluateCell(cell, valuesById, scope, input.inputs ?? {}, indexes, errors);
    valuesById.set(cell.id, value);
    scope[cell.address] = value;
    if (cell.name) scope[cell.name] = value;
    for (const reference of cell.references ?? []) scope[reference] = value;
  }

  for (const cell of cells) {
    if (cell.role !== "validation" || !cell.validation) continue;
    const before = errors.length;
    const passed = Boolean(evaluateExpression(cell.validation.condition, scope, cell, errors));
    if (errors.length > before) {
      ruleStates.push(ruleState(cell, "error"));
    } else if (!passed) {
      errors.push(issue(cell, cell.validation.message));
      ruleStates.push(ruleState(cell, "fail"));
    } else {
      ruleStates.push(ruleState(cell, "ok"));
    }
  }

  for (const cell of cells) {
    if (cell.role !== "compliance" || !cell.compliance) continue;
    const before = errors.length;
    const fired = Boolean(evaluateExpression(cell.compliance.condition, scope, cell, errors));
    if (errors.length > before) {
      ruleStates.push(ruleState(cell, "error"));
    } else if (fired) {
      warnings.push(issue(cell, cell.compliance.message));
      ruleStates.push(ruleState(cell, "warn"));
    } else {
      ruleStates.push(ruleState(cell, "ok"));
    }
  }

  const values = valuesRecord(cells, valuesById);
  const outputs: Record<string, CellValue> = {};
  for (const cell of cells) {
    if (!cell.surfaced) continue;
    outputs[cell.name ?? cell.address] = valuesById.get(cell.id) ?? null;
  }

  return {
    valid: errors.length === 0,
    values,
    outputs,
    executionOrder: sort.order,
    errors,
    warnings,
    ruleStates,
  };
}

export function executeWorkbookEngine(input: WorkbookEngineInput): WorkbookEngineResult {
  const duplicateNames = duplicateWorkbookNames(input.sheets);
  const duplicateNameSet = new Set(duplicateNames);
  const cellMeta = new Map<string, { sheetId: string; sheetName: string; original: EngineCell; internal: EngineCell }>();
  const sheetLookup = buildSheetLookup(input.sheets);
  const internalCells: EngineCell[] = [];

  for (const [sheetIndex, sheet] of input.sheets.entries()) {
    for (const original of sheet.cells) {
      const internalAddress = scopedAddress(sheetIndex, original.address);
      const name = original.name && !duplicateNameSet.has(original.name) ? original.name : null;
      const internal: EngineCell = {
        ...original,
        id: scopedCellId(sheet.id, original.id),
        address: internalAddress,
        name,
        formula: original.formula ? transformWorkbookExpression(original.formula, sheet, sheetIndex, sheetLookup) : original.formula,
        references: [...(original.references ?? []), `${sheet.name}!${original.address}`],
      };

      if (original.lookup) {
        internal.lookup = {
          ...original.lookup,
          inputMap: Object.fromEntries(
            Object.entries(original.lookup.inputMap).map(([column, reference]) => [
              column,
              transformWorkbookReference(reference, sheet, sheetIndex, sheetLookup),
            ]),
          ),
        };
      }

      if (original.validation) {
        internal.validation = {
          ...original.validation,
          condition: transformWorkbookExpression(original.validation.condition, sheet, sheetIndex, sheetLookup),
        };
      }

      if (original.compliance) {
        internal.compliance = {
          ...original.compliance,
          condition: transformWorkbookExpression(original.compliance.condition, sheet, sheetIndex, sheetLookup),
        };
      }

      cellMeta.set(internal.id, { sheetId: sheet.id, sheetName: sheet.name, original, internal });
      internalCells.push(internal);
    }
  }

  const duplicateIssues = duplicateNameIssues(input.sheets, duplicateNames);
  const result = executeEngine({
    cells: internalCells,
    inputs: transformWorkbookInputs(input.inputs ?? {}, input.sheets, sheetLookup),
  });
  const remappedErrors = [...remapWorkbookIssues(result.errors, cellMeta), ...duplicateIssues];
  const remappedWarnings = remapWorkbookIssues(result.warnings, cellMeta);
  const remappedRuleStates = result.ruleStates.map((rule) => {
    const meta = cellMeta.get(rule.cellId);
    return {
      ...rule,
      address: meta ? meta.original.address : rule.address,
      name: meta ? meta.original.name : rule.name,
    };
  });
  const sheetResults = buildWorkbookSheetResults(input.sheets, result, cellMeta, remappedErrors, remappedWarnings, remappedRuleStates);
  const outputs: Record<string, CellValue> = {};

  for (const sheetResult of sheetResults) {
    for (const [key, value] of Object.entries(sheetResult.result.outputs)) {
      outputs[`${sheetResult.sheetName}!${key}`] = value;
    }
  }

  return {
    valid: result.valid && duplicateIssues.length === 0,
    outputs,
    errors: remappedErrors,
    warnings: remappedWarnings,
    ruleStates: remappedRuleStates,
    sheetResults,
  };
}

function buildIndexes(cells: EngineCell[]) {
  const byId = new Map<string, EngineCell>();
  const byReference = new Map<string, EngineCell>();

  for (const cell of cells) {
    byId.set(cell.id, cell);
    byReference.set(cell.address, cell);
    if (cell.name) byReference.set(cell.name, cell);
    for (const reference of cell.references ?? []) byReference.set(reference, cell);
  }

  return { byId, byReference };
}

function referencesForCell(cell: EngineCell): string[] {
  const expressions: string[] = [];
  if (cell.formula) expressions.push(cell.formula);
  if (cell.validation?.condition) expressions.push(cell.validation.condition);
  if (cell.compliance?.condition) expressions.push(cell.compliance.condition);

  const refs = new Set<string>();
  for (const expression of expressions) {
    for (const ref of referencesForExpression(expression)) refs.add(ref);
  }

  if (cell.lookup) {
    for (const ref of Object.values(cell.lookup.inputMap)) {
      for (const lookupRef of referencesForLookupInput(ref)) refs.add(lookupRef);
    }
  }

  return [...refs];
}

function referencesForLookupInput(referenceOrExpression: string): string[] {
  if (!referenceOrExpression.includes("&")) return [referenceOrExpression];
  const refs = new Set<string>();

  for (const part of splitConcatenationExpression(referenceOrExpression)) {
    const trimmed = part.trim();
    if (!trimmed || (trimmed.startsWith('"') && trimmed.endsWith('"'))) continue;
    for (const ref of referencesForExpression(trimmed)) refs.add(ref);
  }

  return [...refs];
}

function referencesForExpression(expression: string): string[] {
  const refs = new Set<string>();
  for (const ref of referencesForLookupCalls(expression)) refs.add(ref);

  const expressionWithoutLookups = replaceLookupCallsWithLiteral(expression, "0");
  for (const ref of referencesForMathExpression(expressionWithoutLookups)) refs.add(ref);

  return [...refs];
}

function referencesForMathExpression(expression: string): string[] {
  let node: MathNode;
  try {
    node = parseExpression(expression);
  } catch {
    return [];
  }
  const refs = new Set<string>();

  node.traverse((child) => {
    if (child.type !== "SymbolNode") return;
    const name = "name" in child ? String(child.name) : "";
    if (!name || ALLOWED_FUNCTIONS.has(name) || ["true", "false", "TRUE", "FALSE"].includes(name)) return;
    refs.add(name);
  });

  return [...refs];
}

function referencesForLookupCalls(expression: string): string[] {
  const refs = new Set<string>();

  for (const call of findLookupFunctionCalls(expression)) {
    if (call.name === "VLOOKUP") {
      for (const ref of referencesForFormulaPart(call.args[0] ?? "")) refs.add(ref);
      for (const ref of referencesForRangeArgument(call.args[1] ?? "")) refs.add(ref);
    }

    if (call.name === "XLOOKUP") {
      for (const ref of referencesForFormulaPart(call.args[0] ?? "")) refs.add(ref);
      for (const ref of referencesForRangeArgument(call.args[1] ?? "")) refs.add(ref);
      for (const ref of referencesForRangeArgument(call.args[2] ?? "")) refs.add(ref);
      if (call.args[3]) {
        for (const ref of referencesForFormulaPart(call.args[3])) refs.add(ref);
      }
    }
  }

  return [...refs];
}

function referencesForFormulaPart(expression: string): string[] {
  if (expression.includes("&")) return referencesForLookupInput(expression);
  return referencesForMathExpression(expression);
}

function referencesForRangeArgument(expression: string): string[] {
  const range = normalizeRangeReference(expression);
  if (!range) return referencesForFormulaPart(expression);
  return expandReferenceRange(range);
}

function evaluateCell(
  cell: EngineCell,
  valuesById: Map<string, CellValue>,
  scope: Record<string, CellValue>,
  inputs: Record<string, CellValue>,
  indexes: ReturnType<typeof buildIndexes>,
  errors: EngineIssue[],
): CellValue {
  const inputOverride = inputs[cell.id] ?? inputs[cell.name ?? ""] ?? inputs[cell.address];

  if (cell.role === "input") {
    return inputOverride ?? cell.value ?? null;
  }

  if (cell.role === "lookup" || (cell.role === "action" && cell.lookup)) {
    if (!cell.lookup) {
      errors.push(issue(cell, "Lookup cell is missing its lookup table definition."));
      return null;
    }
    return evaluateLookup(cell, valuesById, scope, indexes, errors);
  }

  if (cell.formula && (cell.role === "formula" || cell.role === "output" || cell.role === "action")) {
    return evaluateExpression(cell.formula, scope, cell, errors);
  }

  return cell.value ?? null;
}

function evaluateLookup(
  cell: EngineCell,
  valuesById: Map<string, CellValue>,
  scope: Record<string, CellValue>,
  indexes: ReturnType<typeof buildIndexes>,
  errors: EngineIssue[],
): CellValue {
  const lookup = cell.lookup;
  if (!lookup) return null;

  const matched = lookup.rows.find((row) => {
    return Object.entries(lookup.inputMap).every(([column, ref]) => {
      const actual = lookupInputValue(ref, valuesById, scope, indexes, cell, errors);
      return row[column] === actual;
    });
  });

  if (!matched) {
    const criteria = Object.entries(lookup.inputMap)
      .map(([column, ref]) => {
        const actual = lookupInputValue(ref, valuesById, scope, indexes, cell, errors);
        return `${column}=${String(actual ?? "")}`;
      })
      .join(", ");
    errors.push(issue(cell, `Lookup table did not find a matching row${criteria ? ` for ${criteria}` : ""}.`));
    return null;
  }

  return matched[lookup.outputColumn] ?? null;
}

function lookupInputValue(
  referenceOrExpression: string,
  valuesById: Map<string, CellValue>,
  scope: Record<string, CellValue>,
  indexes: ReturnType<typeof buildIndexes>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue | undefined {
  const source = indexes.byReference.get(referenceOrExpression);
  if (source) return valuesById.get(source.id);
  if (!referenceOrExpression.includes("&")) return undefined;
  return evaluateConcatenationExpression(referenceOrExpression, scope, cell, errors);
}

function evaluateConcatenationExpression(
  expression: string,
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue | undefined {
  const parts = splitConcatenationExpression(expression);
  if (parts.length <= 1) return undefined;

  return parts.map((part) => {
    const trimmed = part.trim();
    if (trimmed.startsWith('"') && trimmed.endsWith('"')) return trimmed.slice(1, -1).replace(/""/g, '"');
    if (trimmed in scope) return String(scope[trimmed] ?? "");
    const value = evaluateExpression(trimmed, scope, cell, errors);
    return String(value ?? "");
  }).join("");
}

function splitConcatenationExpression(expression: string): string[] {
  const parts: string[] = [];
  let current = "";
  let depth = 0;
  let inString = false;

  for (const char of expression) {
    if (char === '"') {
      inString = !inString;
      current += char;
      continue;
    }
    if (!inString && char === "(") depth += 1;
    if (!inString && char === ")") depth -= 1;
    if (!inString && depth === 0 && char === "&") {
      parts.push(current);
      current = "";
      continue;
    }
    current += char;
  }

  parts.push(current);
  return parts;
}

function evaluateExpression(
  expression: string,
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue {
  try {
    const expressionWithLookupValues = replaceLookupCallsWithValues(expression, scope, cell, errors);
    const node = parseExpression(expressionWithLookupValues);
    const expressionScope = { ...scope };
    for (const ref of referencesForExpression(expressionWithLookupValues)) {
      if ((isCellReference(ref) || isScopedCellReference(ref)) && !(ref in expressionScope)) expressionScope[ref] = 0;
    }
    const result = node.evaluate(expressionScope);
    return normalizeValue(result);
  } catch (error) {
    errors.push(issue(cell, `Formula error: ${(error as Error).message}`));
    return null;
  }
}

function replaceLookupCallsWithValues(
  expression: string,
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): string {
  const calls = findLookupFunctionCalls(expression);
  if (calls.length === 0) return expression;

  let next = expression;
  for (const call of [...calls].reverse()) {
    const value = evaluateLookupFunction(call.name, call.args, scope, cell, errors);
    next = `${next.slice(0, call.start)}${literalForFormula(value)}${next.slice(call.end + 1)}`;
  }
  return next;
}

function replaceLookupCallsWithLiteral(expression: string, literal: string): string {
  const calls = findLookupFunctionCalls(expression);
  if (calls.length === 0) return expression;

  let next = expression;
  for (const call of [...calls].reverse()) {
    next = `${next.slice(0, call.start)}${literal}${next.slice(call.end + 1)}`;
  }
  return next;
}

function evaluateLookupFunction(
  name: string,
  args: string[],
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue {
  if (name === "VLOOKUP") return evaluateVlookupFunction(args, scope, cell, errors);
  if (name === "XLOOKUP") return evaluateXlookupFunction(args, scope, cell, errors);
  return null;
}

function evaluateVlookupFunction(
  args: string[],
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue {
  if (args.length < 3) throw new Error("VLOOKUP requires lookup value, table range, and output column.");
  const lookupValue = evaluateLookupArgument(args[0], scope, cell, errors);
  const table = rangeValues(args[1], scope);
  const outputColumn = Math.trunc(toNumber(evaluateLookupArgument(args[2], scope, cell, errors)) ?? 0);
  const rangeLookup = args[3] ? evaluateLookupArgument(args[3], scope, cell, errors) : true;

  if (rangeLookup !== false && rangeLookup !== 0 && String(rangeLookup).toLowerCase() !== "false") {
    throw new Error("VLOOKUP approximate matching is not supported yet; use FALSE for exact match.");
  }

  if (outputColumn < 1) throw new Error("VLOOKUP output column must be 1 or greater.");

  for (const row of table) {
    if (lookupMatches(row[0], lookupValue)) return row[outputColumn - 1] ?? null;
  }

  throw new Error(`VLOOKUP did not find a matching row for ${String(lookupValue ?? "")}.`);
}

function evaluateXlookupFunction(
  args: string[],
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue {
  if (args.length < 3) throw new Error("XLOOKUP requires lookup value, lookup range, and return range.");
  const lookupValue = evaluateLookupArgument(args[0], scope, cell, errors);
  const lookupValues = rangeValues(args[1], scope).flat();
  const returnValues = rangeValues(args[2], scope).flat();
  const ifNotFound = args[3];
  const matchMode = args[4] ? evaluateLookupArgument(args[4], scope, cell, errors) : 0;

  if (matchMode !== 0 && String(matchMode).toLowerCase() !== "0") {
    throw new Error("XLOOKUP only supports exact match mode 0 right now.");
  }

  for (let index = 0; index < lookupValues.length; index += 1) {
    if (lookupMatches(lookupValues[index], lookupValue)) return returnValues[index] ?? null;
  }

  if (ifNotFound !== undefined && ifNotFound !== "") return evaluateLookupArgument(ifNotFound, scope, cell, errors);
  throw new Error(`XLOOKUP did not find a matching row for ${String(lookupValue ?? "")}.`);
}

function evaluateLookupArgument(
  expression: string,
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue {
  const trimmed = expression.trim();
  if (/^true$/i.test(trimmed)) return true;
  if (/^false$/i.test(trimmed)) return false;
  if (expression.includes("&")) return evaluateConcatenationExpression(expression, scope, cell, errors) ?? null;
  return evaluateExpression(expression, scope, cell, errors);
}

function rangeValues(expression: string, scope: Record<string, CellValue>): CellValue[][] {
  const range = normalizeRangeReference(expression);
  if (!range) throw new Error(`Lookup range "${expression}" is not a supported cell range.`);
  const refs = expandReferenceRange(range);
  const bounds = referenceRangeBounds(range);
  if (!bounds) return [];

  const rows: CellValue[][] = [];
  let index = 0;
  for (let row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
    const values: CellValue[] = [];
    for (let column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
      values.push(scope[refs[index]] ?? null);
      index += 1;
    }
    rows.push(values);
  }
  return rows;
}

function lookupMatches(left: CellValue, right: CellValue): boolean {
  if (left === right) return true;
  if (left === null || right === null) return false;
  return String(left) === String(right);
}

function literalForFormula(value: CellValue): string {
  if (value === null) return "null";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "null";
  if (typeof value === "boolean") return value ? "true" : "false";
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

function findLookupFunctionCalls(expression: string): Array<{ name: string; args: string[]; start: number; end: number }> {
  const calls: Array<{ name: string; args: string[]; start: number; end: number }> = [];
  const pattern = /\b(VLOOKUP|XLOOKUP)\s*\(/gi;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(expression))) {
    const openIndex = match.index + match[0].lastIndexOf("(");
    const closeIndex = matchingParenIndex(expression, openIndex);
    if (closeIndex === null) continue;
    calls.push({
      name: match[1].toUpperCase(),
      args: splitFormulaArguments(expression.slice(openIndex + 1, closeIndex)),
      start: match.index,
      end: closeIndex,
    });
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

function parseExpression(expression: string): MathNode {
  const normalized = normalizeExpression(expression);
  const node = math.parse(normalized);

  node.traverse((child) => {
    if (DISALLOWED_NODE_TYPES.has(child.type)) {
      throw new Error(`${child.type} is not allowed in Quoin formulas.`);
    }

    if (child.type === "FunctionNode") {
      const fn = (child as unknown as { fn?: { name?: string } }).fn;
      const name = fn?.name ? String(fn.name) : "";
      if (!ALLOWED_FUNCTIONS.has(name)) {
        throw new Error(`Function "${name}" is not allowed in Quoin formulas.`);
      }
    }
  });

  return node;
}

function sumValues(values: unknown[]): number {
  return numericValues(values).reduce((total, value) => total + value, 0);
}

function averageValues(values: unknown[]): number {
  const numbers = numericValues(values);
  if (numbers.length === 0) return 0;
  return sumValues(numbers) / numbers.length;
}

function numericValues(values: unknown[]): number[] {
  return values.flatMap((value) => {
    if (Array.isArray(value)) return numericValues(value);
    const number = toNumber(value);
    return number === null ? [] : [number];
  });
}

function toNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function normalizeExpression(expression: string): string {
  const withoutEquals = expression.trim().startsWith("=") ? expression.trim().slice(1) : expression;
  return expandCellRanges(withoutEquals);
}

function expandCellRanges(expression: string): string {
  return expression.replace(
    /\b((?:__sheet\d+_)?[A-Z]+[1-9]\d*)\s*:\s*((?:__sheet\d+_)?[A-Z]+[1-9]\d*)\b/gi,
    (_match, start: string, end: string) => {
      return expandReferenceRange(`${start}:${end}`).join(", ");
    },
  );
}

function expandRange(start: string, end: string): string[] {
  const startRef = parseCellReference(start);
  const endRef = parseCellReference(end);
  if (!startRef || !endRef) return [start, end];

  const firstColumn = Math.min(startRef.column, endRef.column);
  const lastColumn = Math.max(startRef.column, endRef.column);
  const firstRow = Math.min(startRef.row, endRef.row);
  const lastRow = Math.max(startRef.row, endRef.row);
  const references: string[] = [];

  for (let row = firstRow; row <= lastRow; row++) {
    for (let column = firstColumn; column <= lastColumn; column++) {
      references.push(`${columnName(column)}${row}`);
    }
  }

  return references;
}

function normalizeRangeReference(expression: string): string | null {
  const normalized = expression.trim().replace(/^=/, "").replace(/\$/g, "");
  const match = /^((?:__sheet\d+_)?[A-Z]+[1-9]\d*)\s*:\s*((?:__sheet\d+_)?[A-Z]+[1-9]\d*)$/i.exec(normalized);
  if (!match) return null;
  return `${match[1].toUpperCase()}:${match[2].toUpperCase()}`;
}

function expandReferenceRange(range: string): string[] {
  const bounds = referenceRangeBounds(range);
  if (!bounds) return [];
  const refs: string[] = [];

  for (let row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
    for (let column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
      refs.push(`${bounds.scopePrefix}${columnName(column)}${row}`);
    }
  }

  return refs;
}

function referenceRangeBounds(range: string): {
  scopePrefix: string;
  firstColumn: number;
  lastColumn: number;
  firstRow: number;
  lastRow: number;
} | null {
  const normalized = normalizeRangeReference(range);
  if (!normalized) return null;
  const [start, end] = normalized.split(":");
  const startRef = parseAnyCellReference(start);
  const endRef = parseAnyCellReference(end);
  if (!startRef || !endRef || startRef.scopePrefix !== endRef.scopePrefix) return null;

  return {
    scopePrefix: startRef.scopePrefix,
    firstColumn: Math.min(startRef.column, endRef.column),
    lastColumn: Math.max(startRef.column, endRef.column),
    firstRow: Math.min(startRef.row, endRef.row),
    lastRow: Math.max(startRef.row, endRef.row),
  };
}

function parseCellReference(reference: string): { column: number; row: number } | null {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(reference);
  if (!match) return null;
  return {
    column: columnNumber(match[1]),
    row: Number(match[2]),
  };
}

function parseAnyCellReference(reference: string): { scopePrefix: string; column: number; row: number } | null {
  const scoped = /^(__sheet\d+_)([A-Z]+)([1-9]\d*)$/i.exec(reference);
  if (scoped) {
    return {
      scopePrefix: scoped[1].toLowerCase(),
      column: columnNumber(scoped[2].toUpperCase()),
      row: Number(scoped[3]),
    };
  }

  const plain = parseCellReference(reference.toUpperCase());
  return plain ? { scopePrefix: "", column: plain.column, row: plain.row } : null;
}

function isCellReference(reference: string): boolean {
  return parseCellReference(reference.toUpperCase()) !== null;
}

function isScopedCellReference(reference: string): boolean {
  return /^__sheet\d+_[A-Z]+[1-9]\d*$/.test(reference);
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

function normalizeValue(value: unknown): CellValue {
  if (value === null || value === undefined) return null;
  if (typeof value === "number" || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "object" && "valueOf" in value) {
    const primitive = (value as { valueOf: () => unknown }).valueOf();
    if (typeof primitive === "number" || typeof primitive === "string" || typeof primitive === "boolean") {
      return primitive;
    }
  }
  return String(value);
}

function topologicalSort(cells: EngineCell[], dependencies: Map<string, Set<string>>) {
  const visited = new Set<string>();
  const visiting = new Set<string>();
  const cycleIds = new Set<string>();
  const order: string[] = [];

  const visit = (id: string) => {
    if (visiting.has(id)) {
      cycleIds.add(id);
      return;
    }
    if (visited.has(id)) return;

    visiting.add(id);
    for (const dep of dependencies.get(id) ?? []) visit(dep);
    visiting.delete(id);
    visited.add(id);
    order.push(id);
  };

  for (const cell of cells) visit(cell.id);

  return { order, cycleIds };
}

function valuesRecord(cells: EngineCell[], valuesById: Map<string, CellValue>): Record<string, CellValue> {
  const values: Record<string, CellValue> = {};
  for (const cell of cells) {
    values[cell.address] = valuesById.get(cell.id) ?? null;
    if (cell.name) values[cell.name] = valuesById.get(cell.id) ?? null;
  }
  return values;
}

function emptyResult(
  valid: boolean,
  cells: EngineCell[],
  executionOrder: string[],
  errors: EngineIssue[],
  warnings: EngineIssue[],
  ruleStates: RuleState[],
): EngineResult {
  return {
    valid,
    values: {},
    outputs: Object.fromEntries(cells.filter((cell) => cell.surfaced).map((cell) => [cell.name ?? cell.address, null])),
    executionOrder,
    errors,
    warnings,
    ruleStates,
  };
}

function duplicateWorkbookNames(sheets: WorkbookEngineSheet[]): string[] {
  const counts = new Map<string, number>();

  for (const sheet of sheets) {
    for (const cell of sheet.cells) {
      if (!cell.name) continue;
      counts.set(cell.name, (counts.get(cell.name) ?? 0) + 1);
    }
  }

  return [...counts.entries()].filter(([, count]) => count > 1).map(([name]) => name);
}

function duplicateNameIssues(sheets: WorkbookEngineSheet[], duplicateNames: string[]): EngineIssue[] {
  const duplicateNameSet = new Set(duplicateNames);
  const issues: EngineIssue[] = [];

  for (const sheet of sheets) {
    for (const cell of sheet.cells) {
      if (!cell.name || !duplicateNameSet.has(cell.name)) continue;
      issues.push({
        cellId: scopedCellId(sheet.id, cell.id),
        address: cell.address,
        name: cell.name,
        message: `Duplicate Smart Cell name "${cell.name}" appears on more than one Sheet. Rename one before using workbook-level formulas.`,
      });
    }
  }

  return issues;
}

function buildSheetLookup(sheets: WorkbookEngineSheet[]) {
  const byName = new Map<string, { sheet: WorkbookEngineSheet; index: number }>();

  for (const [index, sheet] of sheets.entries()) {
    byName.set(normalizeSheetName(sheet.name), { sheet, index });
  }

  return byName;
}

function scopedCellId(sheetId: string, cellId: string): string {
  return `${sheetId}!${cellId}`;
}

function scopedAddress(sheetIndex: number, address: string): string {
  return `__sheet${sheetIndex}_${normalizeAddress(address)}`;
}

function normalizeAddress(address: string): string {
  return address.replace(/\$/g, "").toUpperCase();
}

function normalizeSheetName(name: string): string {
  return name.trim().toLowerCase();
}

function transformWorkbookInputs(
  inputs: Record<string, CellValue>,
  sheets: WorkbookEngineSheet[],
  sheetLookup: Map<string, { sheet: WorkbookEngineSheet; index: number }>,
): Record<string, CellValue> {
  const transformed: Record<string, CellValue> = {};

  for (const [key, value] of Object.entries(inputs)) {
    transformed[key] = value;
    for (const sheet of sheets) {
      const sheetIndex = sheetLookup.get(normalizeSheetName(sheet.name))?.index ?? 0;
      transformed[transformWorkbookReference(key, sheet, sheetIndex, sheetLookup)] = value;
    }
  }

  return transformed;
}

function transformWorkbookReference(
  reference: string,
  currentSheet: WorkbookEngineSheet,
  currentSheetIndex: number,
  sheetLookup: Map<string, { sheet: WorkbookEngineSheet; index: number }>,
): string {
  return transformWorkbookExpression(reference, currentSheet, currentSheetIndex, sheetLookup);
}

function transformWorkbookExpression(
  expression: string,
  currentSheet: WorkbookEngineSheet,
  currentSheetIndex: number,
  sheetLookup: Map<string, { sheet: WorkbookEngineSheet; index: number }>,
): string {
  let next = expression;

  next = next.replace(
    /'((?:[^']|'')+)'\!\$?([A-Z]+)\$?([1-9]\d*)\s*:\s*\$?([A-Z]+)\$?([1-9]\d*)/gi,
    (match, sheetName: string, startColumn: string, startRow: string, endColumn: string, endRow: string) => {
      return scopedRangeForSheet(unescapeSheetName(sheetName), `${startColumn}${startRow}`, `${endColumn}${endRow}`, sheetLookup) ?? match;
    },
  );

  next = next.replace(
    /\b([A-Za-z_][A-Za-z0-9_]*)\!\$?([A-Z]+)\$?([1-9]\d*)\s*:\s*\$?([A-Z]+)\$?([1-9]\d*)/g,
    (match, sheetName: string, startColumn: string, startRow: string, endColumn: string, endRow: string) => {
      return scopedRangeForSheet(sheetName, `${startColumn}${startRow}`, `${endColumn}${endRow}`, sheetLookup) ?? match;
    },
  );

  next = next.replace(
    /'((?:[^']|'')+)'\!\$?([A-Z]+)\$?([1-9]\d*)/gi,
    (match, sheetName: string, column: string, row: string) => {
      return scopedAddressForSheet(unescapeSheetName(sheetName), `${column}${row}`, sheetLookup) ?? match;
    },
  );

  next = next.replace(
    /\b([A-Za-z_][A-Za-z0-9_]*)\!\$?([A-Z]+)\$?([1-9]\d*)/g,
    (match, sheetName: string, column: string, row: string) => {
      return scopedAddressForSheet(sheetName, `${column}${row}`, sheetLookup) ?? match;
    },
  );

  next = next.replace(
    /\b\$?([A-Z]+)\$?([1-9]\d*)\s*:\s*\$?([A-Z]+)\$?([1-9]\d*)\b/g,
    (_match, startColumn: string, startRow: string, endColumn: string, endRow: string) => {
      return scopedRange(currentSheetIndex, `${startColumn}${startRow}`, `${endColumn}${endRow}`);
    },
  );

  next = next.replace(/\b\$?([A-Z]+)\$?([1-9]\d*)\b/g, (_match, column: string, row: string) => {
    return scopedAddress(currentSheetIndex, `${column}${row}`);
  });

  return next;
}

function scopedAddressForSheet(
  sheetName: string,
  address: string,
  sheetLookup: Map<string, { sheet: WorkbookEngineSheet; index: number }>,
): string | null {
  const matched = sheetLookup.get(normalizeSheetName(sheetName));
  if (!matched) return null;
  return scopedAddress(matched.index, address);
}

function scopedRangeForSheet(
  sheetName: string,
  start: string,
  end: string,
  sheetLookup: Map<string, { sheet: WorkbookEngineSheet; index: number }>,
): string | null {
  const matched = sheetLookup.get(normalizeSheetName(sheetName));
  if (!matched) return null;
  return scopedRange(matched.index, start, end);
}

function scopedRange(sheetIndex: number, start: string, end: string): string {
  return `${scopedAddress(sheetIndex, start)}:${scopedAddress(sheetIndex, end)}`;
}

function unescapeSheetName(name: string): string {
  return name.replace(/''/g, "'");
}

function remapWorkbookIssues(
  issues: EngineIssue[],
  cellMeta: Map<string, { sheetId: string; sheetName: string; original: EngineCell; internal: EngineCell }>,
): EngineIssue[] {
  return issues.map((item) => {
    const meta = cellMeta.get(item.cellId);
    if (!meta) return item;
    return {
      ...item,
      address: meta.original.address,
      name: meta.original.name,
    };
  });
}

function buildWorkbookSheetResults(
  sheets: WorkbookEngineSheet[],
  result: EngineResult,
  cellMeta: Map<string, { sheetId: string; sheetName: string; original: EngineCell; internal: EngineCell }>,
  errors: EngineIssue[],
  warnings: EngineIssue[],
  ruleStates: RuleState[],
): WorkbookEngineSheetResult[] {
  return sheets.map((sheet) => {
    const sheetCells = [...cellMeta.values()].filter((meta) => meta.sheetId === sheet.id);
    const sheetCellIds = new Set(sheetCells.map((meta) => meta.internal.id));
    const values: Record<string, CellValue> = {};
    const outputs: Record<string, CellValue> = {};

    for (const meta of sheetCells) {
      const value = result.values[meta.internal.name ?? meta.internal.address] ?? result.values[meta.internal.address] ?? null;
      values[meta.original.address] = value;
      if (meta.original.name) values[meta.original.name] = value;
      if (meta.original.surfaced) outputs[meta.original.name ?? meta.original.address] = value;
    }

    return {
      sheetId: sheet.id,
      sheetName: sheet.name,
      result: {
        valid: result.valid && !errors.some((item) => sheetCellIds.has(item.cellId)),
        values,
        outputs,
        executionOrder: result.executionOrder.filter((id) => sheetCellIds.has(id)),
        errors: errors.filter((item) => sheetCellIds.has(item.cellId)),
        warnings: warnings.filter((item) => sheetCellIds.has(item.cellId)),
        ruleStates: ruleStates.filter((item) => sheetCellIds.has(item.cellId)),
      },
    };
  });
}

function issue(cell: EngineCell, message: string): EngineIssue {
  return {
    cellId: cell.id,
    address: cell.address,
    name: cell.name,
    message,
  };
}

function ruleState(cell: EngineCell, state: RuleState["state"]): RuleState {
  return {
    cellId: cell.id,
    address: cell.address,
    name: cell.name,
    state,
  };
}
