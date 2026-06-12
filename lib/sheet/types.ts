import type { CellValue, SmartCellRole, SmartCellType } from "@/lib/engine";

export interface LookupConfig {
  inputColumn: string;
  inputReference: string;
  inputs?: Array<{ column: string; reference: string }>;
  outputColumn: string;
  rows: Array<Record<string, CellValue>>;
}

export interface GridCell {
  address: string;
  entry: string;
  name: string;
  label: string;
  role: SmartCellRole;
  type: SmartCellType;
  inputOptions: string[];
  surfaced: boolean;
  annotation: string;
  ruleMessage: string;
  lookup?: LookupConfig;
}

export interface LocalConfiguration {
  id: string;
  name: string;
  cells: Record<string, GridCell>;
  columnCount?: number;
  rowCount?: number;
  updatedAt: string;
}

export interface SheetSnapshot {
  cells: Record<string, GridCell>;
  columnCount: number;
  rowCount: number;
}
