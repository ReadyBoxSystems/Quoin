import type { GridCell, WorkbookSheet } from "@/lib/sheet/types";

export interface WorkbookState {
  sheets: WorkbookSheet[];
  activeSheetId: string;
}

export function createWorkbookState(sheets: WorkbookSheet[], activeSheetId?: string): WorkbookState {
  const active = sheets.find((sheet) => sheet.id === activeSheetId) ?? sheets[0];
  return { sheets, activeSheetId: active?.id ?? "" };
}

export function getActiveSheet(workbook: WorkbookState): WorkbookSheet | null {
  return workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0] ?? null;
}

export function replaceWorkbook(sheets: WorkbookSheet[], activeSheetId?: string): WorkbookState {
  return createWorkbookState(sheets, activeSheetId);
}

export function switchActiveSheet(workbook: WorkbookState, sheetId: string): WorkbookState {
  if (sheetId === workbook.activeSheetId || !workbook.sheets.some((sheet) => sheet.id === sheetId)) return workbook;
  return { ...workbook, activeSheetId: sheetId };
}

export function updateActiveSheet(
  workbook: WorkbookState,
  update: (sheet: WorkbookSheet) => WorkbookSheet,
): WorkbookState {
  const active = getActiveSheet(workbook);
  if (!active) return workbook;
  const next = update(active);
  if (next === active) return workbook;
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => sheet.id === active.id ? next : sheet),
  };
}

export function updateSheet(
  workbook: WorkbookState,
  sheetId: string,
  update: (sheet: WorkbookSheet) => WorkbookSheet,
): WorkbookState {
  const current = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!current) return workbook;
  const next = update(current);
  if (next === current) return workbook;
  return { ...workbook, sheets: workbook.sheets.map((sheet) => sheet.id === sheetId ? next : sheet) };
}

export function addWorkbookSheet(workbook: WorkbookState, sheet: WorkbookSheet): WorkbookState {
  return { sheets: [...workbook.sheets, sheet], activeSheetId: sheet.id };
}

export interface SheetPatch {
  sheetId: string;
  beforeSheet?: WorkbookSheet;
  afterSheet?: WorkbookSheet;
  beforeCells?: Record<string, GridCell | undefined>;
  afterCells?: Record<string, GridCell | undefined>;
  beforeName?: string;
  afterName?: string;
  beforeColumnCount?: number;
  afterColumnCount?: number;
  beforeRowCount?: number;
  afterRowCount?: number;
}

export interface WorkbookTransaction {
  label: string;
  activeSheetIdBefore: string;
  activeSheetIdAfter: string;
  sheetPatches: SheetPatch[];
}

export function diffWorkbook(
  before: WorkbookState,
  after: WorkbookState,
  label: string,
): WorkbookTransaction | null {
  const beforeById = new Map(before.sheets.map((sheet) => [sheet.id, sheet]));
  const afterById = new Map(after.sheets.map((sheet) => [sheet.id, sheet]));
  const ids = new Set([...beforeById.keys(), ...afterById.keys()]);
  const sheetPatches: SheetPatch[] = [];

  for (const sheetId of ids) {
    const oldSheet = beforeById.get(sheetId);
    const newSheet = afterById.get(sheetId);
    if (!oldSheet || !newSheet) {
      sheetPatches.push({ sheetId, beforeSheet: oldSheet, afterSheet: newSheet });
      continue;
    }

    const addresses = new Set([...Object.keys(oldSheet.cells), ...Object.keys(newSheet.cells)]);
    const beforeCells: Record<string, GridCell | undefined> = {};
    const afterCells: Record<string, GridCell | undefined> = {};
    for (const address of addresses) {
      if (oldSheet.cells[address] === newSheet.cells[address]) continue;
      if (JSON.stringify(oldSheet.cells[address]) === JSON.stringify(newSheet.cells[address])) continue;
      beforeCells[address] = oldSheet.cells[address];
      afterCells[address] = newSheet.cells[address];
    }
    const cellsChanged = Object.keys(beforeCells).length > 0;
    const metadataChanged = oldSheet.name !== newSheet.name
      || oldSheet.columnCount !== newSheet.columnCount
      || oldSheet.rowCount !== newSheet.rowCount;
    if (!cellsChanged && !metadataChanged) continue;
    sheetPatches.push({
      sheetId,
      beforeCells: cellsChanged ? beforeCells : undefined,
      afterCells: cellsChanged ? afterCells : undefined,
      beforeName: oldSheet.name,
      afterName: newSheet.name,
      beforeColumnCount: oldSheet.columnCount,
      afterColumnCount: newSheet.columnCount,
      beforeRowCount: oldSheet.rowCount,
      afterRowCount: newSheet.rowCount,
    });
  }

  if (sheetPatches.length === 0 && before.activeSheetId === after.activeSheetId) return null;
  return { label, activeSheetIdBefore: before.activeSheetId, activeSheetIdAfter: after.activeSheetId, sheetPatches };
}

export function applyWorkbookTransaction(
  workbook: WorkbookState,
  transaction: WorkbookTransaction,
  direction: "undo" | "redo",
): WorkbookState {
  let sheets = [...workbook.sheets];
  for (const patch of transaction.sheetPatches) {
    const wholeSheet = direction === "undo" ? patch.beforeSheet : patch.afterSheet;
    const oppositeWholeSheet = direction === "undo" ? patch.afterSheet : patch.beforeSheet;
    if (wholeSheet || oppositeWholeSheet) {
      sheets = sheets.filter((sheet) => sheet.id !== patch.sheetId);
      if (wholeSheet) sheets.push(wholeSheet);
      continue;
    }

    const index = sheets.findIndex((sheet) => sheet.id === patch.sheetId);
    if (index < 0) continue;
    const sourceCells = direction === "undo" ? patch.beforeCells : patch.afterCells;
    const cells = { ...sheets[index].cells };
    for (const [address, cell] of Object.entries(sourceCells ?? {})) {
      if (cell) cells[address] = cell;
      else delete cells[address];
    }
    sheets[index] = {
      ...sheets[index],
      cells,
      name: (direction === "undo" ? patch.beforeName : patch.afterName) ?? sheets[index].name,
      columnCount: (direction === "undo" ? patch.beforeColumnCount : patch.afterColumnCount) ?? sheets[index].columnCount,
      rowCount: (direction === "undo" ? patch.beforeRowCount : patch.afterRowCount) ?? sheets[index].rowCount,
    };
  }
  return createWorkbookState(sheets, direction === "undo" ? transaction.activeSheetIdBefore : transaction.activeSheetIdAfter);
}
