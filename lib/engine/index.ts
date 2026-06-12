import { all, create, type MathJsInstance, type MathNode } from "mathjs";
import type { CellValue, EngineCell, EngineInput, EngineIssue, EngineResult, RuleState } from "./types";

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
        if (isCellReference(ref)) continue;
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

function buildIndexes(cells: EngineCell[]) {
  const byId = new Map<string, EngineCell>();
  const byReference = new Map<string, EngineCell>();

  for (const cell of cells) {
    byId.set(cell.id, cell);
    byReference.set(cell.address, cell);
    if (cell.name) byReference.set(cell.name, cell);
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
    for (const ref of Object.values(cell.lookup.inputMap)) refs.add(ref);
  }

  return [...refs];
}

function referencesForExpression(expression: string): string[] {
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
    if (!name || ALLOWED_FUNCTIONS.has(name) || name === "true" || name === "false") return;
    refs.add(name);
  });

  return [...refs];
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
    return evaluateLookup(cell, valuesById, indexes, errors);
  }

  if (cell.formula && (cell.role === "formula" || cell.role === "output" || cell.role === "action")) {
    return evaluateExpression(cell.formula, scope, cell, errors);
  }

  return cell.value ?? null;
}

function evaluateLookup(
  cell: EngineCell,
  valuesById: Map<string, CellValue>,
  indexes: ReturnType<typeof buildIndexes>,
  errors: EngineIssue[],
): CellValue {
  const lookup = cell.lookup;
  if (!lookup) return null;

  const matched = lookup.rows.find((row) => {
    return Object.entries(lookup.inputMap).every(([column, ref]) => {
      const source = indexes.byReference.get(ref);
      const actual = source ? valuesById.get(source.id) : undefined;
      return row[column] === actual;
    });
  });

  if (!matched) {
    const criteria = Object.entries(lookup.inputMap)
      .map(([column, ref]) => {
        const source = indexes.byReference.get(ref);
        const actual = source ? valuesById.get(source.id) : undefined;
        return `${column}=${String(actual ?? "")}`;
      })
      .join(", ");
    errors.push(issue(cell, `Lookup table did not find a matching row${criteria ? ` for ${criteria}` : ""}.`));
    return null;
  }

  return matched[lookup.outputColumn] ?? null;
}

function evaluateExpression(
  expression: string,
  scope: Record<string, CellValue>,
  cell: EngineCell,
  errors: EngineIssue[],
): CellValue {
  try {
    const node = parseExpression(expression);
    const expressionScope = { ...scope };
    for (const ref of referencesForExpression(expression)) {
      if (isCellReference(ref) && !(ref in expressionScope)) expressionScope[ref] = 0;
    }
    const result = node.evaluate(expressionScope);
    return normalizeValue(result);
  } catch (error) {
    errors.push(issue(cell, `Formula error: ${(error as Error).message}`));
    return null;
  }
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
  return expression.replace(/\b([A-Z]+[1-9]\d*)\s*:\s*([A-Z]+[1-9]\d*)\b/gi, (_match, start: string, end: string) => {
    return expandRange(start.toUpperCase(), end.toUpperCase()).join(", ");
  });
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

function parseCellReference(reference: string): { column: number; row: number } | null {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(reference);
  if (!match) return null;
  return {
    column: columnNumber(match[1]),
    row: Number(match[2]),
  };
}

function isCellReference(reference: string): boolean {
  return parseCellReference(reference.toUpperCase()) !== null;
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
