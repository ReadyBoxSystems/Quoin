import { executeEngine, type EngineCell } from "./index";

let failed = 0;

function test(name: string, fn: () => void) {
  try {
    fn();
    console.log(`  ok - ${name}`);
  } catch (error) {
    failed++;
    console.log(`  FAIL - ${name}`);
    console.log(`    ${(error as Error).message}`);
  }
}

function assertEq<T>(actual: T, expected: T, message: string) {
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  if (a !== b) throw new Error(`${message}\n  expected: ${b}\n  actual:   ${a}`);
}

function cell(overrides: Partial<EngineCell> & Pick<EngineCell, "id" | "address" | "role">): EngineCell {
  return {
    type: "number",
    surfaced: false,
    ...overrides,
  };
}

console.log("executeEngine");

test("evaluates named references", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "a", address: "A1", name: "wall_height", role: "input", value: 10 }),
      cell({ id: "b", address: "B1", name: "wall_width", role: "input", value: 12 }),
      cell({ id: "c", address: "C1", name: "wall_area", role: "output", formula: "wall_height * wall_width", surfaced: true }),
    ],
  });

  assertEq(result.valid, true, "engine should be valid");
  assertEq(result.outputs.wall_area, 120, "named formula should evaluate");
});

test("evaluates coordinate references", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "a", address: "A1", role: "input", value: 5 }),
      cell({ id: "b", address: "B1", role: "input", value: 7 }),
      cell({ id: "c", address: "C1", name: "total", role: "output", formula: "A1 + B1", surfaced: true }),
    ],
  });

  assertEq(result.outputs.total, 12, "coordinate formula should evaluate");
});

test("evaluates Excel-style range formulas", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "a", address: "A1", role: "input", value: 5 }),
      cell({ id: "b", address: "A2", role: "input", value: 7 }),
      cell({ id: "c", address: "A3", role: "input", value: 9 }),
      cell({ id: "total", address: "B1", name: "total", role: "output", formula: "SUM(A1:A3)", surfaced: true }),
      cell({ id: "average", address: "B2", name: "average", role: "output", formula: "AVERAGE(A1:A3)", surfaced: true }),
    ],
  });

  assertEq(result.valid, true, "range formulas should be valid");
  assertEq(result.outputs.total, 21, "SUM should evaluate a cell range");
  assertEq(result.outputs.average, 7, "AVERAGE should evaluate a cell range");
});

test("evaluates rectangular ranges", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "a", address: "A1", role: "input", value: 1 }),
      cell({ id: "b", address: "B1", role: "input", value: 2 }),
      cell({ id: "c", address: "A2", role: "input", value: 3 }),
      cell({ id: "d", address: "B2", role: "input", value: 4 }),
      cell({ id: "total", address: "C1", name: "total", role: "output", formula: "SUM(A1:B2)", surfaced: true }),
    ],
  });

  assertEq(result.valid, true, "rectangular range should be valid");
  assertEq(result.outputs.total, 10, "SUM should evaluate every cell in a rectangular range");
});

test("allows blank coordinate cells inside ranges", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "a", address: "A1", role: "input", value: 5 }),
      cell({ id: "c", address: "A3", role: "input", value: 9 }),
      cell({ id: "total", address: "B1", name: "total", role: "output", formula: "SUM(A1:A3)", surfaced: true }),
    ],
  });

  assertEq(result.valid, true, "blank cells in a range should not make the formula invalid");
  assertEq(result.outputs.total, 14, "SUM should ignore blank coordinate cells in a range");
});

test("evaluates IF formulas", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "span", address: "A1", name: "span", role: "input", value: 14 }),
      cell({
        id: "band",
        address: "B1",
        name: "load_band",
        role: "output",
        type: "text",
        formula: 'IF(span > 12, "heavy", "standard")',
        surfaced: true,
      }),
    ],
  });

  assertEq(result.valid, true, "IF formula should be valid");
  assertEq(result.outputs.load_band, "heavy", "IF should return the matching branch");
});

test("surfaces action outputs", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "span", address: "A1", name: "span", role: "input", value: 16 }),
      cell({ id: "action", address: "B1", name: "shop_action", role: "action", type: "text", value: "Verify bearing.", surfaced: true }),
    ],
  });

  assertEq(result.valid, true, "action output should be valid");
  assertEq(result.outputs.shop_action, "Verify bearing.", "action output should be surfaced");
});

test("detects missing references", () => {
  const result = executeEngine({
    cells: [cell({ id: "a", address: "A1", name: "bad", role: "output", formula: "missing_value + 1", surfaced: true })],
  });

  assertEq(result.valid, false, "engine should be invalid");
  assertEq(result.errors[0].message, 'Missing reference "missing_value".', "missing reference should be reported");
});

test("detects circular references", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "a", address: "A1", name: "a", role: "formula", formula: "b + 1" }),
      cell({ id: "b", address: "B1", name: "b", role: "formula", formula: "a + 1" }),
    ],
  });

  assertEq(result.valid, false, "engine should be invalid");
  if (!result.errors.some((error) => error.message.includes("circular"))) {
    throw new Error("expected circular dependency error");
  }
});

test("executes lookup tables", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "species", address: "A1", name: "species", role: "input", type: "text", value: "SPF" }),
      cell({ id: "span", address: "B1", name: "span", role: "input", value: 12 }),
      cell({
        id: "beam",
        address: "C1",
        name: "beam_size",
        role: "lookup",
        type: "text",
        surfaced: true,
        lookup: {
          inputMap: { species: "species", span: "span" },
          outputColumn: "beam",
          rows: [
            { species: "SPF", span: 10, beam: "2x8" },
            { species: "SPF", span: 12, beam: "2x10" },
          ],
        },
      }),
    ],
  });

  assertEq(result.valid, true, "lookup should be valid");
  assertEq(result.outputs.beam_size, "2x10", "lookup should return matching output");
});

test("preserves resolved values when a lookup misses", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "span", address: "A1", name: "span", role: "input", value: 9 }),
      cell({
        id: "beam",
        address: "B1",
        name: "beam_size",
        role: "lookup",
        type: "text",
        surfaced: true,
        lookup: {
          inputMap: { span: "span" },
          outputColumn: "beam",
          rows: [{ span: 12, beam: "2x10" }],
        },
      }),
    ],
  });

  assertEq(result.valid, false, "lookup miss should make result invalid");
  assertEq(result.values.span, 9, "input value should remain available");
  assertEq(result.outputs.beam_size, null, "failed lookup output should be null");
});

test("blocks failed validation and emits compliance warnings", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "span", address: "A1", name: "span", role: "input", value: 18 }),
      cell({
        id: "validation",
        address: "B1",
        name: "span_limit",
        role: "validation",
        validation: { condition: "span <= 20", message: "Span exceeds absolute limit." },
      }),
      cell({
        id: "warning",
        address: "C1",
        name: "engineer_review",
        role: "compliance",
        compliance: { condition: "span > 16", message: "Engineer review recommended." },
      }),
    ],
  });

  assertEq(result.valid, true, "validation should pass");
  assertEq(result.warnings.length, 1, "compliance warning should fire");
  assertEq(result.ruleStates, [
    { cellId: "validation", address: "B1", name: "span_limit", state: "ok" },
    { cellId: "warning", address: "C1", name: "engineer_review", state: "warn" },
  ], "rule states should show validation ok and compliance warn");
});

test("marks failed validation as failed while preserving computed context", () => {
  const result = executeEngine({
    cells: [
      cell({ id: "span", address: "A1", name: "span", role: "input", value: 21 }),
      cell({
        id: "validation",
        address: "B1",
        name: "span_limit",
        role: "validation",
        validation: { condition: "span <= 20", message: "No spans over 20 feet." },
      }),
      cell({
        id: "warning",
        address: "C1",
        name: "engineer_review",
        role: "compliance",
        compliance: { condition: "span > 11", message: "Engineer review recommended." },
      }),
    ],
  });

  assertEq(result.valid, false, "validation failure should make result invalid");
  assertEq(result.values.span, 21, "input value should remain available");
  assertEq(result.ruleStates, [
    { cellId: "validation", address: "B1", name: "span_limit", state: "fail" },
    { cellId: "warning", address: "C1", name: "engineer_review", state: "warn" },
  ], "rule states should show validation fail and compliance warn");
});

test("is deterministic", () => {
  const input = {
    cells: [
      cell({ id: "a", address: "A1", name: "x", role: "input", value: 3 }),
      cell({ id: "b", address: "B1", name: "y", role: "output", formula: "x * 2", surfaced: true }),
    ],
  };

  assertEq(executeEngine(input), executeEngine(input), "same input should produce same result");
});

if (failed > 0) {
  console.log(`\n${failed} test(s) failed`);
  process.exit(1);
}

console.log("\nAll tests passed");
