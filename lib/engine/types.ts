export type CellValue = number | string | boolean | null;

export type SmartCellRole =
  | "input"
  | "formula"
  | "output"
  | "action"
  | "lookup"
  | "validation"
  | "compliance";

export type SmartCellType = "number" | "text" | "boolean";

export interface LookupTableDefinition {
  inputMap: Record<string, string>;
  outputColumn: string;
  rows: Array<Record<string, CellValue>>;
}

export interface RuleDefinition {
  condition: string;
  message: string;
}

export interface EngineCell {
  id: string;
  address: string;
  references?: string[];
  name?: string | null;
  role: SmartCellRole;
  type: SmartCellType;
  value?: CellValue;
  formula?: string | null;
  annotation?: string | null;
  surfaced?: boolean;
  lookup?: LookupTableDefinition;
  validation?: RuleDefinition;
  compliance?: RuleDefinition;
}

export interface EngineInput {
  cells: EngineCell[];
  inputs?: Record<string, CellValue>;
}

export interface WorkbookEngineSheet {
  id: string;
  name: string;
  cells: EngineCell[];
}

export interface WorkbookEngineInput {
  sheets: WorkbookEngineSheet[];
  inputs?: Record<string, CellValue>;
}

export interface EngineIssue {
  cellId: string;
  address: string;
  name?: string | null;
  message: string;
}

export interface RuleState {
  cellId: string;
  address: string;
  name?: string | null;
  state: "ok" | "warn" | "fail" | "error";
}

export interface EngineResult {
  valid: boolean;
  values: Record<string, CellValue>;
  outputs: Record<string, CellValue>;
  executionOrder: string[];
  errors: EngineIssue[];
  warnings: EngineIssue[];
  ruleStates: RuleState[];
}

export interface WorkbookEngineSheetResult {
  sheetId: string;
  sheetName: string;
  result: EngineResult;
}

export interface WorkbookEngineResult {
  valid: boolean;
  outputs: Record<string, CellValue>;
  errors: EngineIssue[];
  warnings: EngineIssue[];
  ruleStates: RuleState[];
  sheetResults: WorkbookEngineSheetResult[];
}
