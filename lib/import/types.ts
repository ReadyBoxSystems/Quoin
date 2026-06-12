import type { CellValue } from "@/lib/engine/types";

export type ImportedCellValue = CellValue | Date;

export type ImportedCellKind = "blank" | "value" | "formula";

export interface ImportedCell {
  address: string;
  kind: ImportedCellKind;
  value?: ImportedCellValue;
  formula?: string;
}

export interface ImportedSheetDimensions {
  rowCount: number;
  columnCount: number;
}

export interface ImportedSheet {
  id: string;
  name: string;
  dimensions: ImportedSheetDimensions;
  cells: ImportedCell[];
}

export type ImportedNameKind = "singleCell" | "range" | "formula" | "external" | "unknown";

export interface ImportedName {
  name: string;
  kind: ImportedNameKind;
  reference: string;
  sheetName?: string;
}

export type ImportReviewSeverity = "info" | "warning" | "error";

export interface ImportReviewItem {
  severity: ImportReviewSeverity;
  message: string;
  sheetName?: string;
  address?: string;
  formula?: string;
  name?: string;
}

export interface ImportedWorkbook {
  fileName: string;
  sheets: ImportedSheet[];
  names: ImportedName[];
  reviewItems: ImportReviewItem[];
}

export interface ImportedSheetSummary {
  sheetName: string;
  cellCount: number;
  formulaCount: number;
  namedCellCount: number;
  namedRangeCount: number;
  reviewItemCount: number;
}

export interface ImportSummary {
  fileName: string;
  sheetCount: number;
  namesCount: number;
  sheets: ImportedSheetSummary[];
  reviewItems: ImportReviewItem[];
}
