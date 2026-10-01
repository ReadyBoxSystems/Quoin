"use client";

import { useEffect, useMemo, useRef, useState, type ChangeEvent, type ComponentProps, type ReactNode } from "react";
import { CommandMenu } from "@/components/command-menu";
import { SheetStrip } from "@/components/sheet-strip";
import {
  executeEngine,
  executeWorkbookEngine,
  type CellValue,
  type EngineCell,
  type SmartCellRole,
  type SmartCellType,
  type WorkbookEngineResult,
} from "@/lib/engine";
import { convertImportedSheetToQuoin } from "@/lib/import/convert";
import { parseFormula, quoteSheetName as quoteFormulaSheetName, rewriteFormula } from "@/lib/formula/model";
import type { ImportedName, ImportedWorkbook, ImportReviewItem } from "@/lib/import/types";
import type { GridCell, InputControl, LocalConfiguration, LookupConfig, SmartCellVisibilityCondition, WorkbookSheet } from "@/lib/sheet/types";
import {
  addWorkbookSheet,
  applyWorkbookTransaction,
  createWorkbookState,
  diffWorkbook,
  getActiveSheet,
  replaceWorkbook,
  switchActiveSheet,
  updateActiveSheet,
  type WorkbookState,
  type WorkbookTransaction,
} from "@/lib/workbook/state";

const STORAGE_KEY = "quoin.gridSheet.v2";
const LEGACY_CONFIG_STORAGE_KEY = "quoin.configurations.v1";
const CONFIG_STORAGE_KEY = "quoin.configurations.v2";
const ACTIVE_CONFIG_KEY = "quoin.activeConfiguration.v1";
const defaultColumnCount = 16;
const defaultRowCount = 30;
const historyLimit = 50;
const coreRoleOptions: SmartCellRole[] = ["input", "formula", "output"];
const advancedRoleOptions: SmartCellRole[] = ["lookup", "action", "validation", "compliance"];
const typeOptions: SmartCellType[] = ["number", "text", "boolean"];
const inputControlOptions: InputControl[] = ["freeText", "dropdown", "checkbox"];

type RunnerOverrides = Record<string, string>;

interface DependencyItem {
  address: string;
  label: string;
  reference: string;
}

interface DependencySummary {
  dependencies: DependencyItem[];
  dependents: DependencyItem[];
}

interface RunnerSheetContext {
  sheetId: string;
  sheetName: string;
  cells: Record<string, GridCell>;
  displayValues: Record<string, CellValue>;
  surfacedCells: GridCell[];
  result: ReturnType<typeof executeEngine>;
  validationStates: Array<{ address: string; state: string; name?: string | null }>;
}

interface VisibilityControlOption {
  name: string;
  label: string;
  sheetName: string;
}

interface ReferenceOption {
  address: string;
  key: string;
  reference: string;
  value: CellValue;
}

const beamLookup: LookupConfig = {
  inputColumn: "design_span",
  inputReference: "design_span",
  inputs: [
    { column: "design_span", reference: "design_span" },
    { column: "load_band", reference: "load_band" },
    { column: "story_condition", reference: "story_condition" },
  ],
  outputColumn: "beam",
  rows: [
    { design_span: 10, load_band: "standard", story_condition: "top_floor", beam: "2x10 SPF" },
    { design_span: 12, load_band: "standard", story_condition: "top_floor", beam: "2x12 SPF" },
    { design_span: 14, load_band: "standard", story_condition: "first_floor", beam: "9.25 LVL" },
    { design_span: 16, load_band: "standard", story_condition: "first_floor", beam: "11.875 LVL" },
    { design_span: 14, load_band: "heavy", story_condition: "first_floor", beam: "11.875 LVL" },
    { design_span: 16, load_band: "heavy", story_condition: "first_floor", beam: "14 LVL" },
    { design_span: 12, load_band: "standard", story_condition: "dropped_header", beam: "9.25 LVL" },
    { design_span: 14, load_band: "standard", story_condition: "dropped_header", beam: "11.875 LVL" },
  ],
};

const memberLookup: LookupConfig = {
  inputColumn: "design_span",
  inputReference: "design_span",
  inputs: [
    { column: "design_span", reference: "design_span" },
    { column: "load_band", reference: "load_band" },
    { column: "story_condition", reference: "story_condition" },
  ],
  outputColumn: "members",
  rows: [
    { design_span: 10, load_band: "standard", story_condition: "top_floor", members: 2 },
    { design_span: 12, load_band: "standard", story_condition: "top_floor", members: 2 },
    { design_span: 14, load_band: "standard", story_condition: "first_floor", members: 3 },
    { design_span: 16, load_band: "standard", story_condition: "first_floor", members: 4 },
    { design_span: 14, load_band: "heavy", story_condition: "first_floor", members: 4 },
    { design_span: 16, load_band: "heavy", story_condition: "first_floor", members: 4 },
    { design_span: 12, load_band: "standard", story_condition: "dropped_header", members: 2 },
    { design_span: 14, load_band: "standard", story_condition: "dropped_header", members: 3 },
  ],
};

const shopNoteLookup: LookupConfig = {
  inputColumn: "design_span",
  inputReference: "design_span",
  inputs: [
    { column: "design_span", reference: "design_span" },
    { column: "load_band", reference: "load_band" },
    { column: "story_condition", reference: "story_condition" },
  ],
  outputColumn: "note",
  rows: [
    { design_span: 10, load_band: "standard", story_condition: "top_floor", note: "Standard fastening schedule." },
    { design_span: 12, load_band: "standard", story_condition: "top_floor", note: "Standard fastening schedule." },
    { design_span: 14, load_band: "standard", story_condition: "first_floor", note: "Add fastening review note." },
    { design_span: 16, load_band: "standard", story_condition: "first_floor", note: "Add fastening review note and verify bearing." },
    { design_span: 14, load_band: "heavy", story_condition: "first_floor", note: "Use heavy-load fastening review." },
    { design_span: 16, load_band: "heavy", story_condition: "first_floor", note: "Escalate fastening and bearing review." },
    { design_span: 12, load_band: "standard", story_condition: "dropped_header", note: "Dropped header acceptable in this demo table." },
    { design_span: 14, load_band: "standard", story_condition: "dropped_header", note: "Dropped header requires fastening review." },
  ],
};

const starterLookup = beamLookup;

const initialCells: Record<string, GridCell> = {
  A1: makeCell("A1", "Beam Selection Demo", "text"),
  A2: makeCell("A2", "Design span", "text"),
  B2: makeCell("B2", "14", "number", {
    name: "design_span",
    label: "Design Span (ft)",
    role: "input",
    surfaced: true,
    annotation: "Span taken from the drawing. Demo lookup supports 10, 12, 14, and 16 ft.",
  }),
  A3: makeCell("A3", "Design PLF", "text"),
  B3: makeCell("B3", "650", "number", {
    name: "design_plf",
    label: "Design PLF",
    role: "input",
    surfaced: true,
    annotation: "Demo line load used to show calculated context.",
  }),
  A4: makeCell("A4", "Story condition", "text"),
  B4: makeCell("B4", "first_floor", "text", {
    name: "story_condition",
    label: "Story Condition",
    role: "input",
    surfaced: true,
    annotation: "Simple demo condition label for the runner.",
    inputOptions: ["top_floor", "first_floor", "dropped_header"],
  }),
  A5: makeCell("A5", "Load band", "text"),
  B5: makeCell("B5", "standard", "text", {
    name: "load_band",
    label: "Load Band",
    role: "input",
    surfaced: true,
    annotation: "Controlled demo category used by the lookup table.",
    inputOptions: ["standard", "heavy"],
  }),
  A6: makeCell("A6", "Total line load", "text"),
  B6: makeCell("B6", "=design_span * design_plf", "number", {
    name: "total_line_load",
    label: "Total Line Load",
    role: "output",
    surfaced: true,
    annotation: "Calculated context from span and PLF.",
  }),
  A7: makeCell("A7", "Recommended beam", "text"),
  B7: makeCell("B7", "=LOOKUP(design_span)", "text", {
    name: "recommended_beam",
    label: "Recommended Beam",
    role: "lookup",
    surfaced: true,
    annotation: "Demo recommendation from a small fake lookup table.",
    lookup: beamLookup,
  }),
  A8: makeCell("A8", "Member count", "text"),
  B8: makeCell("B8", "=LOOKUP(design_span)", "number", {
    name: "member_count",
    label: "Member Count",
    role: "lookup",
    surfaced: true,
    annotation: "Demo member count from a small fake lookup table.",
    lookup: memberLookup,
  }),
  A9: makeCell("A9", "Shop note", "text"),
  B9: makeCell("B9", "=LOOKUP(design_span)", "text", {
    name: "shop_note",
    label: "Shop Note",
    role: "action",
    surfaced: true,
    annotation: "Demo shop note tied to the selected span.",
    lookup: shopNoteLookup,
  }),
  A11: makeCell("A11", "Span limit check", "text"),
  B11: makeCell("B11", "=design_span <= 16", "boolean", {
    name: "span_limit",
    label: "Span Limit Check",
    role: "validation",
    annotation: "Demo validation that keeps the span inside the available table.",
    ruleMessage: "Span is beyond the demo table. Escalate for engineering review.",
  }),
  A12: makeCell("A12", "Engineering review", "text"),
  B12: makeCell("B12", "=design_span > 14", "boolean", {
    name: "engineer_review",
    label: "Engineer Review",
    role: "compliance",
    annotation: "Demo warning for longer spans.",
    ruleMessage: "Engineering review recommended for spans over 14 ft in this demo table.",
  }),
};

export function VariableSheet() {
  const [workbook, setWorkbook] = useState<WorkbookState>(() => createWorkbookState([
    makeWorkbookSheet("Sheet 1", initialCells, defaultColumnCount, defaultRowCount),
  ]));
  const [configurations, setConfigurations] = useState<LocalConfiguration[]>([]);
  const [activeConfigId, setActiveConfigId] = useState("");
  const [configName, setConfigName] = useState("Demo - Beam Selection");
  const [isDirty, setIsDirty] = useState(false);
  const [selectedAddress, setSelectedAddress] = useState("B2");
  const [editingAddress, setEditingAddress] = useState<string | null>(null);
  const [isFormulaBarActive, setIsFormulaBarActive] = useState(false);
  const [draftEntry, setDraftEntry] = useState("");
  const [isLoaded, setIsLoaded] = useState(false);
  const [activeView, setActiveView] = useState<"sheet" | "runner" | "help">("sheet");
  const [activeReferenceIndex, setActiveReferenceIndex] = useState(0);
  const [undoStack, setUndoStack] = useState<WorkbookTransaction[]>([]);
  const [redoStack, setRedoStack] = useState<WorkbookTransaction[]>([]);
  const [copiedAddress, setCopiedAddress] = useState<string | null>(null);
  const [pendingImport, setPendingImport] = useState<ImportedWorkbook | null>(null);
  const [selectedImportSheetId, setSelectedImportSheetId] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [importMessage, setImportMessage] = useState("");
  const [importError, setImportError] = useState("");
  const [runnerOverrides, setRunnerOverrides] = useState<RunnerOverrides>({});
  const [isInspectorOpen, setIsInspectorOpen] = useState(true);
  const [openMenu, setOpenMenu] = useState<string | null>(null);
  const [gridWindowStart, setGridWindowStart] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const configSelectRef = useRef<HTMLSelectElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const cellRefs = useRef<Record<string, HTMLDivElement | null>>({});

  const sheets = workbook.sheets;
  const activeSheetId = workbook.activeSheetId;
  const activeSheet = useMemo(() => getActiveSheet(workbook), [workbook]);
  const cells = activeSheet?.cells ?? {};
  const columnCount = activeSheet?.columnCount ?? defaultColumnCount;
  const rowCount = activeSheet?.rowCount ?? defaultRowCount;
  const columns = useMemo(() => makeColumns(columnCount), [columnCount]);
  const selectedCell = getCell(cells, selectedAddress);
  const engineCells = useMemo(() => toEngineCells(cells), [cells]);
  const visibleSheets = sheets;
  const workbookEngineSheets = useMemo(
    () => visibleSheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      cells: toEngineCells(sheet.cells),
    })),
    [visibleSheets],
  );
  const workbookResult = useMemo(() => executeWorkbookEngine({ sheets: workbookEngineSheets }), [workbookEngineSheets]);
  const runnerVisibleSheets = useMemo(
    () => applyRunnerOverridesToSheets(visibleSheets, runnerOverrides),
    [runnerOverrides, visibleSheets],
  );
  const runnerWorkbookEngineSheets = useMemo(
    () => runnerVisibleSheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      cells: toEngineCells(sheet.cells),
    })),
    [runnerVisibleSheets],
  );
  const hasRunnerOverrides = Object.keys(runnerOverrides).length > 0;
  const runnerWorkbookResult = useMemo(
    () => hasRunnerOverrides ? executeWorkbookEngine({ sheets: runnerWorkbookEngineSheets }) : workbookResult,
    [hasRunnerOverrides, runnerWorkbookEngineSheets, workbookResult],
  );
  const result = useMemo(
    () => workbookResult.sheetResults.find((item) => item.sheetId === activeSheet?.id)?.result ?? executeEngine({ cells: engineCells }),
    [activeSheet?.id, engineCells, workbookResult],
  );
  const runnerSheets = useMemo(
    () => buildRunnerSheetContexts(runnerVisibleSheets, runnerWorkbookResult),
    [runnerVisibleSheets, runnerWorkbookResult],
  );
  const visibleRunnerKeys = useMemo(
    () => new Set(runnerSheets.flatMap((sheet) => sheet.surfacedCells.filter((cell) => cell.role === "input").map((cell) => runnerOverrideKey(sheet.sheetId, cell.address)))),
    [runnerSheets],
  );
  const visibilityControls = useMemo(
    () => buildVisibilityControlOptions(visibleSheets),
    [visibleSheets],
  );
  const ruleStateMap = useMemo(() => new Map(result.ruleStates.map((rule) => [rule.address, rule.state])), [result.ruleStates]);
  const displayValues = useMemo(
    () => buildDisplayValues(cells, result.values, result.errors, ruleStateMap, columns, rowCount),
    [cells, columns, result.errors, result.values, rowCount, ruleStateMap],
  );
  const columnWidths = useMemo(() => buildColumnWidths(cells, displayValues, columns, rowCount), [cells, columns, displayValues, rowCount]);
  const referenceOptions = useMemo(
    () => buildWorkbookReferenceOptions(visibleSheets, workbookResult, activeSheetId),
    [activeSheetId, visibleSheets, workbookResult],
  );
  const filteredReferenceOptions = useMemo(() => {
    const query = getReferenceQuery(draftEntry);
    const normalized = query.toLowerCase();
    if (!normalized) return referenceOptions;

    return referenceOptions.filter((option) => {
      return option.reference.toLowerCase().includes(normalized) || option.address.toLowerCase().includes(normalized);
    });
  }, [draftEntry, referenceOptions]);
  const visibleReferenceOptions = useMemo(() => {
    const editingKey = activeSheet ? referenceOptionKey(activeSheet.id, editingAddress ?? "") : "";
    return filteredReferenceOptions.filter((option) => option.key !== editingKey).slice(0, 10);
  }, [activeSheet, editingAddress, filteredReferenceOptions]);
  const issueMap = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const issue of [...result.errors, ...result.warnings]) {
      const list = map.get(issue.address) ?? [];
      list.push(issue.message);
      map.set(issue.address, list);
    }
    return map;
  }, [result.errors, result.warnings]);
  const dependencySummary = useMemo(
    () => buildDependencySummary(cells, selectedAddress),
    [cells, selectedAddress],
  );
  useEffect(() => {
    try {
      const storedConfigs = window.localStorage.getItem(CONFIG_STORAGE_KEY)
        ?? window.localStorage.getItem(LEGACY_CONFIG_STORAGE_KEY);
      const activeId = window.localStorage.getItem(ACTIVE_CONFIG_KEY) ?? "";
      const parsedConfigs = storedConfigs ? hydrateConfigurations(JSON.parse(storedConfigs) as LocalConfiguration[]) : [];
      const migratedCells = migrateLegacyCells();
      const nextConfigurations = parsedConfigs.length > 0
        ? parsedConfigs
        : [makeConfiguration("Demo - Beam Selection", migratedCells ?? initialCells)];
      const activeConfig = nextConfigurations.find((configuration) => configuration.id === activeId) ?? nextConfigurations[0];

      setConfigurations(nextConfigurations);
      setActiveConfigId(activeConfig.id);
      setConfigName(activeConfig.name);
      const loadedSheets = activeConfig.sheets ?? [sheetFromConfiguration(activeConfig)];
      setWorkbook(replaceWorkbook(loadedSheets, activeConfig.activeSheetId));
      setUndoStack([]);
      setRedoStack([]);
      setIsDirty(false);
    } finally {
      setIsLoaded(true);
    }
  }, []);

  useEffect(() => {
    if (!isLoaded) return;
    window.localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(configurations));
  }, [configurations, isLoaded]);

  useEffect(() => {
    if (!isLoaded || !activeConfigId) return;
    window.localStorage.setItem(ACTIVE_CONFIG_KEY, activeConfigId);
  }, [activeConfigId, isLoaded]);

  useEffect(() => {
    if (!editingAddress) return;
    if (isFormulaBarActive) return;
    editInputRef.current?.focus();
    editInputRef.current?.select();
  }, [editingAddress, isFormulaBarActive]);

  useEffect(() => {
    setRunnerOverrides((current) => {
      const next = Object.fromEntries(Object.entries(current).filter(([key]) => visibleRunnerKeys.has(key)));
      return Object.keys(next).length === Object.keys(current).length ? current : next;
    });
  }, [visibleRunnerKeys]);

  useEffect(() => {
    if (editingAddress || isFormulaBarActive || activeView !== "sheet") return;
    cellRefs.current[selectedAddress]?.focus();
  }, [activeView, editingAddress, isFormulaBarActive, selectedAddress]);

  useEffect(() => {
    setActiveReferenceIndex(0);
  }, [draftEntry, editingAddress]);

  useEffect(() => {
    setActiveReferenceIndex((current) => Math.min(current, Math.max(0, visibleReferenceOptions.length - 1)));
  }, [visibleReferenceOptions.length]);

  useEffect(() => {
    if (rowCount <= 200) return;
    const selected = parseAddress(selectedAddress);
    if (!selected) return;
    setGridWindowStart((current) => {
      if (selected.row > current && selected.row <= current + 100) return current;
      return Math.max(0, Math.min(rowCount - 100, selected.row - 20));
    });
  }, [rowCount, selectedAddress]);

  useEffect(() => {
    function handleUndoRedo(event: KeyboardEvent) {
      if (event.defaultPrevented) return;
      if (!(event.ctrlKey || event.metaKey)) return;
      const target = event.target as HTMLElement | null;
      if (target?.closest("input, textarea, select")) return;
      if (editingAddress) return;

      if (event.key.toLowerCase() === "z") {
        event.preventDefault();
        if (event.shiftKey) redoCells();
        else undoCells();
      } else if (event.key.toLowerCase() === "y") {
        event.preventDefault();
        redoCells();
      }
    }

    window.addEventListener("keydown", handleUndoRedo);
    return () => window.removeEventListener("keydown", handleUndoRedo);
  });

  function commitWorkbookChange(label: string, transition: (current: WorkbookState) => WorkbookState) {
    setWorkbook((current) => {
      const next = transition(current);
      const transaction = diffWorkbook(current, next, label);
      if (!transaction) return current;
      setUndoStack((history) => [...history.slice(Math.max(0, history.length - historyLimit + 1)), transaction]);
      setRedoStack([]);
      setIsDirty(true);
      return next;
    });
  }

  function applyCellsChange(updater: (current: Record<string, GridCell>) => Record<string, GridCell>, label = "Edit cells") {
    commitWorkbookChange(label, (current) => updateActiveSheet(current, (sheet) => {
      const next = updater(sheet.cells);
      return cellsEqual(sheet.cells, next) ? sheet : { ...sheet, cells: next };
    }));
  }

  function applySheetResize(updater: (current: Record<string, GridCell>) => Record<string, GridCell>, nextColumnCount: number, nextRowCount: number, label: string) {
    commitWorkbookChange(label, (current) => updateActiveSheet(current, (sheet) => ({
      ...sheet,
      cells: updater(sheet.cells),
      columnCount: nextColumnCount,
      rowCount: nextRowCount,
    })));
  }

  function updateCell(address: string, patch: Partial<GridCell>) {
    const existing = getCell(cells, address);
    const oldName = existing.name;
    const nextName = patch.name;
    commitWorkbookChange(`Edit ${address}`, (current) => {
      let next = updateActiveSheet(current, (sheet) => {
        const currentCell = getCell(sheet.cells, address);
        const nextCells = { ...sheet.cells, [address]: applyCellPatch(currentCell, patch) };
        return cellsEqual(sheet.cells, nextCells) ? sheet : { ...sheet, cells: nextCells };
      });
      if (nextName !== undefined && oldName && nextName && oldName !== nextName) {
        next = {
          ...next,
          sheets: next.sheets.map((sheet) => ({
            ...sheet,
            cells: renameVisibilityConditionReferences(sheet.cells, oldName, nextName),
          })),
        };
      }
      return next;
    });
  }

  function updateRunnerCell(sheetId: string, address: string, entry: string) {
    setRunnerOverrides((current) => ({
      ...current,
      [runnerOverrideKey(sheetId, address)]: entry,
    }));
  }

  function clearCell(address: string) {
    applyCellsChange((current) => {
      const next = { ...current };
      delete next[address];
      return next;
    });
  }

  function copyCell(address: string) {
    const cell = getCell(cells, address);
    if (!cell.entry && !cell.name) return;
    setCopiedAddress(address);
  }

  function pasteCopiedCell(targetAddress: string) {
    if (!copiedAddress) return;
    copyCellEntry(copiedAddress, targetAddress);
  }

  function fillDown(address: string) {
    const position = parseAddress(address);
    if (!position || position.row <= 1) return;
    copyCellEntry(`${position.column}${position.row - 1}`, address);
  }

  function copyCellEntry(sourceAddress: string, targetAddress: string) {
    const source = getCell(cells, sourceAddress);
    if (!source.entry) return;

    const sourcePosition = parseAddress(sourceAddress);
    const targetPosition = parseAddress(targetAddress);
    if (!sourcePosition || !targetPosition) return;

    const rowOffset = targetPosition.row - sourcePosition.row;
    const columnOffset = columnNumber(targetPosition.column) - columnNumber(sourcePosition.column);
    const entry = source.entry.trim().startsWith("=")
      ? adjustFormulaReferences(source.entry, rowOffset, columnOffset)
      : source.entry;

    updateCell(targetAddress, { entry, type: source.type });
  }

  function undoCells() {
    if (undoStack.length === 0) return;
    setUndoStack((history) => {
      const previous = history[history.length - 1];
      if (!previous) return history;
      setRedoStack((redoHistory) => [...redoHistory.slice(Math.max(0, redoHistory.length - historyLimit + 1)), previous]);
      setWorkbook((current) => applyWorkbookTransaction(current, previous, "undo"));
      setEditingAddress(null);
      setDraftEntry("");
      setIsDirty(true);
      return history.slice(0, -1);
    });
  }

  function redoCells() {
    if (redoStack.length === 0) return;
    setRedoStack((history) => {
      const next = history[history.length - 1];
      if (!next) return history;

      setUndoStack((undoHistory) => [...undoHistory.slice(Math.max(0, undoHistory.length - historyLimit + 1)), next]);
      setWorkbook((current) => applyWorkbookTransaction(current, next, "redo"));
      setEditingAddress(null);
      setDraftEntry("");
      setIsDirty(true);
      return history.slice(0, -1);
    });
  }

  function startEditing(address: string, replacement?: string) {
    const cell = getCell(cells, address);
    setIsFormulaBarActive(false);
    setSelectedAddress(address);
    setEditingAddress(address);
    setDraftEntry(replacement ?? cell.entry);
  }

  function commitEditing(nextAddress?: string) {
    if (!editingAddress) {
      if (nextAddress) setSelectedAddress(nextAddress);
      setIsFormulaBarActive(false);
      return;
    }
    updateCell(editingAddress, { entry: draftEntry });
    setEditingAddress(null);
    setIsFormulaBarActive(false);
    if (nextAddress) setSelectedAddress(nextAddress);
  }

  function handleCellMouseDown(event: React.MouseEvent<HTMLDivElement>, address: string) {
    if (event.button !== 0) return;
    const target = event.target as HTMLElement | null;
    if (target?.closest("input, select, textarea")) return;

    if (editingAddress && editingAddress !== address) {
      commitEditing(address);
      return;
    }

    setSelectedAddress(address);
  }

  function cancelEditing() {
    setEditingAddress(null);
    setIsFormulaBarActive(false);
    setDraftEntry("");
  }

  function resetRunnerPreview() {
    setRunnerOverrides({});
  }

  function moveSelection(address: string, direction: "up" | "down" | "left" | "right") {
    const position = parseAddress(address);
    if (!position) return address;

    const colIndex = columnNumber(position.column) - 1;
    let nextColIndex = colIndex;
    let nextRow = position.row;

    if (direction === "left") nextColIndex = Math.max(0, colIndex - 1);
    if (direction === "right") nextColIndex = Math.min(columns.length - 1, colIndex + 1);
    if (direction === "up") nextRow = Math.max(1, position.row - 1);
    if (direction === "down") nextRow = Math.min(rowCount, position.row + 1);

    return `${columns[nextColIndex]}${nextRow}`;
  }

  function handleGridKeyDown(event: React.KeyboardEvent<HTMLDivElement>, address: string) {
    const isEditing = editingAddress === address;

    if (!isEditing && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
      event.preventDefault();
      if (event.shiftKey) redoCells();
      else undoCells();
      return;
    }
    if (!isEditing && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "y") {
      event.preventDefault();
      redoCells();
      return;
    }
    if (!isEditing && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      event.preventDefault();
      copyCell(address);
      return;
    }
    if (!isEditing && copiedAddress && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
      event.preventDefault();
      pasteCopiedCell(address);
      return;
    }
    if (!isEditing && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
      event.preventDefault();
      fillDown(address);
      return;
    }

    if (isEditing) {
      if (event.key === "Enter") {
        event.preventDefault();
        commitEditing(moveSelection(address, "down"));
      } else if (event.key === "Tab") {
        event.preventDefault();
        commitEditing(moveSelection(address, event.shiftKey ? "left" : "right"));
      } else if (event.key === "Escape") {
        event.preventDefault();
        cancelEditing();
      }
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      startEditing(address);
      return;
    }
    if (event.key === "F2") {
      event.preventDefault();
      startEditing(address);
      return;
    }
    if (event.key === "Tab") {
      event.preventDefault();
      setSelectedAddress(moveSelection(address, event.shiftKey ? "left" : "right"));
      return;
    }
    if (event.key === "ArrowDown" || event.key === "ArrowUp" || event.key === "ArrowLeft" || event.key === "ArrowRight") {
      event.preventDefault();
      const direction = event.key.replace("Arrow", "").toLowerCase() as "up" | "down" | "left" | "right";
      setSelectedAddress(moveSelection(address, direction));
      return;
    }
    if (event.key === "Backspace" || event.key === "Delete") {
      event.preventDefault();
      clearCell(address);
      return;
    }
    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      startEditing(address, event.key);
    }
  }

  function handleCellClick(address: string) {
    setSelectedAddress(address);
  }

  function handleFormulaBarKeyDown(event: React.KeyboardEvent<HTMLInputElement>) {
    const canUseReferencePopup = draftEntry.trim().startsWith("=") && visibleReferenceOptions.length > 0;

    if (canUseReferencePopup && event.key === "ArrowDown") {
      event.preventDefault();
      setActiveReferenceIndex((current) => Math.min(visibleReferenceOptions.length - 1, current + 1));
      return;
    }
    if (canUseReferencePopup && event.key === "ArrowUp") {
      event.preventDefault();
      setActiveReferenceIndex((current) => Math.max(0, current - 1));
      return;
    }
    if (canUseReferencePopup && event.key === "Tab") {
      event.preventDefault();
      insertReference(visibleReferenceOptions[activeReferenceIndex]?.reference ?? visibleReferenceOptions[0].reference);
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      commitEditing(moveSelection(selectedAddress, "down"));
    } else if (event.key === "Tab") {
      event.preventDefault();
      commitEditing(moveSelection(selectedAddress, event.shiftKey ? "left" : "right"));
    } else if (event.key === "Escape") {
      event.preventDefault();
      cancelEditing();
      cellRefs.current[selectedAddress]?.focus();
    }
  }

  function handleGridPaste(event: React.ClipboardEvent<HTMLDivElement>, address: string) {
    if (editingAddress) return;
    const text = event.clipboardData.getData("text/plain");
    if (!text) return;

    const rows = text.replace(/\r/g, "").split("\n").filter((row, index, allRows) => row !== "" || index < allRows.length - 1);
    if (rows.length === 0) return;

    event.preventDefault();
    const origin = parseAddress(address);
    if (!origin) return;

    const originColumnIndex = columns.indexOf(origin.column);
    applyCellsChange((current) => {
      const next = { ...current };
      rows.forEach((row, rowOffset) => {
        const targetRow = origin.row + rowOffset;
        if (targetRow > rowCount) return;

        row.split("\t").forEach((entry, columnOffset) => {
          const targetColumn = columns[originColumnIndex + columnOffset];
          if (!targetColumn) return;

          const targetAddress = `${targetColumn}${targetRow}`;
          const existing = getCell(next, targetAddress);
          next[targetAddress] = applyCellPatch(existing, { entry });
        });
      });
      return next;
    });
  }

  function resetSheet() {
    if (!confirmDiscardUnsaved()) return;
    commitWorkbookChange("Load demo", (current) => updateActiveSheet(current, (sheet) => ({
      ...sheet,
      cells: initialCells,
      columnCount: defaultColumnCount,
      rowCount: defaultRowCount,
    })));
    setSelectedAddress("B2");
    setEditingAddress(null);
    setRunnerOverrides({});
  }

  function clearSheet() {
    if (!confirmDiscardUnsaved()) return;
    applyCellsChange(() => ({}));
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
    setRunnerOverrides({});
  }

  function currentWorkbookSheets(): WorkbookSheet[] {
    return sheets.length === 0 ? [makeWorkbookSheet("Sheet 1", {}, defaultColumnCount, defaultRowCount)] : sheets;
  }

  function saveConfiguration() {
    const name = configName.trim() || "Untitled Configuration";
    const updatedAt = new Date().toISOString();
    let nextActiveConfigId = activeConfigId;
    const nextSheets = currentWorkbookSheets();
    const nextActiveSheet = nextSheets.find((sheet) => sheet.id === activeSheetId) ?? nextSheets[0];

    setConfigurations((current) => {
      const existing = current.find((configuration) => configuration.id === activeConfigId);
      if (!existing) {
        const created = makeConfiguration(name, nextActiveSheet.cells, nextActiveSheet.columnCount, nextActiveSheet.rowCount, {
          sheets: nextSheets,
          activeSheetId: nextActiveSheet.id,
        });
        nextActiveConfigId = created.id;
        return [...current, created];
      }

      return current.map((configuration) => {
        if (configuration.id !== activeConfigId) return configuration;
        const { cells: _legacyCells, columnCount: _legacyColumns, rowCount: _legacyRows, ...currentConfiguration } = configuration;
        return {
          ...currentConfiguration,
          name,
          activeSheetId: nextActiveSheet.id,
          sheets: nextSheets,
          updatedAt,
        };
      });
    });

    setActiveConfigId(nextActiveConfigId);
    setConfigName(name);
    setIsDirty(false);
  }

  function addRowBelow() {
    const selected = parseAddress(selectedAddress);
    const insertAt = selected ? selected.row + 1 : rowCount + 1;
    applySheetResize((current) => insertRow(current, insertAt), columnCount, Math.max(rowCount + 1, insertAt), `Insert row ${insertAt}`);
    setSelectedAddress(`${selected?.column ?? "A"}${insertAt}`);
    setEditingAddress(null);
  }

  function deleteSelectedRow() {
    const selected = parseAddress(selectedAddress);
    if (!selected || rowCount <= 1) return;
    applySheetResize((current) => deleteRow(current, selected.row), columnCount, Math.max(1, rowCount - 1), `Delete row ${selected.row}`);
    setSelectedAddress(`${selected.column}${Math.min(selected.row, rowCount - 1)}`);
    setEditingAddress(null);
  }

  function addColumnRight() {
    const selected = parseAddress(selectedAddress);
    const selectedColumnNumber = selected ? columnNumber(selected.column) : columnCount;
    const insertAt = selectedColumnNumber + 1;
    applySheetResize((current) => insertColumn(current, insertAt), Math.max(columnCount + 1, insertAt), rowCount, `Insert column ${columnName(insertAt) ?? insertAt}`);
    setSelectedAddress(`${columnName(insertAt) ?? "A"}${selected?.row ?? 1}`);
    setEditingAddress(null);
  }

  function deleteSelectedColumn() {
    const selected = parseAddress(selectedAddress);
    if (!selected || columnCount <= 1) return;
    const deleteAt = columnNumber(selected.column);
    applySheetResize((current) => deleteColumn(current, deleteAt), Math.max(1, columnCount - 1), rowCount, `Delete column ${selected.column}`);
    setSelectedAddress(`${columnName(Math.min(deleteAt, columnCount - 1)) ?? "A"}${selected.row}`);
    setEditingAddress(null);
  }

  function createConfiguration() {
    if (!confirmDiscardUnsaved()) return;
    const firstSheet = makeWorkbookSheet("Sheet 1", {}, defaultColumnCount, defaultRowCount);
    const created = makeConfiguration("Untitled Configuration", firstSheet.cells, firstSheet.columnCount, firstSheet.rowCount, {
      sheets: [firstSheet],
      activeSheetId: firstSheet.id,
    });
    setConfigurations((current) => [...current, created]);
    loadConfiguration(created);
  }

  function duplicateConfiguration() {
    const sourceSheets = currentWorkbookSheets();
    const nextSheets = sourceSheets.map((sheet) => ({ ...sheet, id: makeSheetId() }));
    const activeIndex = Math.max(0, sourceSheets.findIndex((sheet) => sheet.id === activeSheetId));
    const nextActiveSheet = nextSheets[activeIndex] ?? nextSheets[0];
    const created = makeConfiguration(`${configName.trim() || "Configuration"} Copy`, nextActiveSheet.cells, nextActiveSheet.columnCount, nextActiveSheet.rowCount, {
      sheets: nextSheets,
      activeSheetId: nextActiveSheet.id,
    });
    setConfigurations((current) => [...current, created]);
    loadConfiguration(created);
  }

  function deleteConfiguration() {
    const activeConfig = configurations.find((configuration) => configuration.id === activeConfigId);
    const label = activeConfig?.name ?? "this configuration";

    if (!window.confirm(`Delete "${label}"? This only removes the local browser copy.`)) return;

    if (configurations.length <= 1) {
      const created = makeConfiguration("Untitled Configuration", {}, defaultColumnCount, defaultRowCount);
      setConfigurations([created]);
      loadConfiguration(created);
      return;
    }

    const remaining = configurations.filter((configuration) => configuration.id !== activeConfigId);
    setConfigurations(remaining);
    loadConfiguration(remaining[0]);
  }

  function handleConfigurationChange(nextId: string) {
    if (!confirmDiscardUnsaved()) return;
    const next = configurations.find((configuration) => configuration.id === nextId);
    if (next) loadConfiguration(next);
  }

  function loadConfiguration(configuration: LocalConfiguration) {
    const nextSheets = configuration.sheets ?? [sheetFromConfiguration(configuration)];
    const nextActiveSheet = nextSheets.find((sheet) => sheet.id === configuration.activeSheetId) ?? nextSheets[0];
    setActiveConfigId(configuration.id);
    setConfigName(configuration.name);
    setWorkbook(replaceWorkbook(nextSheets, nextActiveSheet.id));
    setUndoStack([]);
    setRedoStack([]);
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
    setRunnerOverrides({});
    setIsDirty(false);
  }

  function switchSheet(nextSheetId: string) {
    if (nextSheetId === activeSheetId) return;
    if (!sheets.some((sheet) => sheet.id === nextSheetId)) return;
    setWorkbook((current) => switchActiveSheet(current, nextSheetId));
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
    setCopiedAddress(null);
  }

  function addSheet() {
    const created = makeWorkbookSheet(`Sheet ${sheets.length + 1}`, {}, defaultColumnCount, defaultRowCount);
    commitWorkbookChange("Add Sheet", (current) => addWorkbookSheet(current, created));
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
  }

  function renameSheet(sheetId: string, name: string) {
    const nextName = name.trimStart();
    const workbookSheets = currentWorkbookSheets();
    const renamedSheet = workbookSheets.find((sheet) => sheet.id === sheetId);
    const finalName = nextName || "Untitled Sheet";
    if (!renamedSheet || renamedSheet.name === finalName) return;

    const nextSheets = workbookSheets.map((sheet) => ({
      ...sheet,
      name: sheet.id === sheetId ? finalName : sheet.name,
      cells: renameSheetReferences(sheet.cells, renamedSheet.name, finalName),
    }));
    const nextActiveSheet = nextSheets.find((sheet) => sheet.id === activeSheetId) ?? nextSheets[0];
    if (nextActiveSheet) commitWorkbookChange("Rename Sheet", () => replaceWorkbook(nextSheets, nextActiveSheet.id));
  }

  function confirmDiscardUnsaved() {
    if (!isDirty) return true;
    return window.confirm("You have unsaved changes. Continue without saving them?");
  }

  function updateLookup(patch: Partial<LookupConfig>) {
    const current = selectedCell.lookup ?? starterLookup;
    updateCell(selectedAddress, { lookup: { ...current, ...patch } });
  }

  function insertReference(reference: string) {
    setDraftEntry((current) => {
      const tokenStart = getReferenceTokenStart(current);
      if (tokenStart !== null) return `${current.slice(0, tokenStart)}${reference}`;

      const needsSpace = current.length > 0 && !/[=\s+\-*/(]$/.test(current);
      return `${current}${needsSpace ? " " : ""}${reference}`;
    });
  }

  async function handleImportFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    setImportMessage("");
    setImportError("");
    setPendingImport(null);

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setImportError("Choose an .xlsx workbook.");
      return;
    }

    setIsImporting(true);
    try {
      const { readExcelWorkbook } = await import("@/lib/import/read");
      const workbook = await readExcelWorkbook(file.name, await file.arrayBuffer());
      if (workbook.sheets.length === 0) {
        setImportError("No worksheets were found in that workbook.");
        return;
      }

      setPendingImport(workbook);
      setSelectedImportSheetId(workbook.sheets[0].id);
      setImportMessage(`Read ${workbook.sheets.length} Sheet${workbook.sheets.length === 1 ? "" : "s"} from ${file.name}.`);
    } catch (error) {
      setImportError(error instanceof Error ? error.message : "Could not read that workbook.");
    } finally {
      setIsImporting(false);
    }
  }

  function confirmImportSheet() {
    if (!pendingImport) return;
    if (!confirmDiscardUnsaved()) return;

    const selectedSheet = pendingImport.sheets.find((sheet) => sheet.id === selectedImportSheetId) ?? pendingImport.sheets[0];
    if (!selectedSheet) {
      setImportError("Choose a worksheet to import.");
      return;
    }

    const convertedSheets = pendingImport.sheets.map((sheet) => {
      const converted = convertImportedSheetToQuoin(sheet, { names: pendingImport.names, workbookSheets: pendingImport.sheets });
      return {
        source: sheet,
        converted,
        workbookSheet: makeWorkbookSheet(sheet.name, converted.cells, converted.columnCount, converted.rowCount),
      };
    });
    const selectedIndex = Math.max(0, pendingImport.sheets.findIndex((sheet) => sheet.id === selectedSheet.id));
    const activeImportedSheet = convertedSheets[selectedIndex] ?? convertedSheets[0];
    const workbookSheets = convertedSheets.map((sheet) => sheet.workbookSheet);
    const configurationName = makeImportedConfigurationName(pendingImport.fileName);
    const created = makeConfiguration(
      configurationName,
      activeImportedSheet.workbookSheet.cells,
      activeImportedSheet.workbookSheet.columnCount,
      activeImportedSheet.workbookSheet.rowCount,
      {
        sheets: workbookSheets,
        activeSheetId: activeImportedSheet.workbookSheet.id,
      },
    );

    setConfigurations((current) => [...current, created]);
    loadConfiguration(created);
    setPendingImport(null);
    setImportError("");
    setImportMessage("");
    setActiveView("sheet");
  }

  function cancelPendingImport() {
    setPendingImport(null);
    setImportError("");
  }

  function openWorkbookIssue() {
    if (!firstStatusIssue) return;
    if (firstStatusIssue.sheetId) {
      setWorkbook((current) => switchActiveSheet(current, firstStatusIssue.sheetId!));
    }
    setSelectedAddress(firstStatusIssue.address);
    setEditingAddress(null);
    setActiveView("sheet");
  }

  const selectedIssues = issueMap.get(selectedAddress) ?? [];
  const firstStatusIssue = workbookResult.errors[0] ?? workbookResult.warnings[0] ?? null;
  const activeSheetHasIssues = result.errors.length > 0;
  const selectedImportSheet = pendingImport?.sheets.find((sheet) => sheet.id === selectedImportSheetId) ?? pendingImport?.sheets[0] ?? null;
  const importReviewItems = pendingImport && selectedImportSheet
    ? pendingImport.reviewItems.concat(importReviewItemsForSheet(pendingImport.names, selectedImportSheet.name))
    : [];

  return (
    <>
      <header className="topbar">
        <div className="productMark">
          <p className="eyebrow">Quoin Core</p>
          <h1>Variable Sheet</h1>
        </div>
        <nav className="menuBar" aria-label="Application commands">
          <CommandMenu label="File" open={openMenu === "File"} onOpenChange={(open) => setOpenMenu((current) => open ? "File" : current === "File" ? null : current)}>
            <button type="button" onClick={createConfiguration}>New configuration</button>
            <button type="button" onClick={() => configSelectRef.current?.focus()}>Choose local configuration</button>
            <button type="button" onClick={saveConfiguration}>Save</button>
            <button type="button" onClick={duplicateConfiguration}>Duplicate</button>
            <button type="button" onClick={() => fileInputRef.current?.click()} disabled={isImporting}>
              {isImporting ? "Importing..." : "Import Workbook"}
            </button>
            <button type="button" className="dangerCommand" onClick={deleteConfiguration}>Delete configuration</button>
          </CommandMenu>
          {activeView === "sheet" && (
            <>
              <CommandMenu label="Edit" open={openMenu === "Edit"} onOpenChange={(open) => setOpenMenu((current) => open ? "Edit" : current === "Edit" ? null : current)}>
                <button type="button" onClick={() => copyCell(selectedAddress)}>Copy Cell</button>
                <button type="button" onClick={() => pasteCopiedCell(selectedAddress)} disabled={!copiedAddress}>Paste Cell</button>
                <button type="button" onClick={() => fillDown(selectedAddress)}>Fill Down</button>
              </CommandMenu>
              <CommandMenu label="Sheet" open={openMenu === "Sheet"} onOpenChange={(open) => setOpenMenu((current) => open ? "Sheet" : current === "Sheet" ? null : current)}>
                <button type="button" onClick={addRowBelow}>Add Row</button>
                <button type="button" onClick={deleteSelectedRow}>Delete Row</button>
                <button type="button" onClick={addColumnRight}>Add Column</button>
                <button type="button" onClick={deleteSelectedColumn}>Delete Column</button>
                <button type="button" className="dangerCommand" onClick={clearSheet}>Clear Sheet</button>
              </CommandMenu>
            </>
          )}
          <CommandMenu label="View" open={openMenu === "View"} onOpenChange={(open) => setOpenMenu((current) => open ? "View" : current === "View" ? null : current)}>
            <button type="button" onClick={() => setActiveView("sheet")}>Sheet</button>
            <button type="button" onClick={() => setActiveView("runner")}>Runner Preview</button>
            {activeView === "sheet" && (
              <button type="button" onClick={() => setIsInspectorOpen((current) => !current)}>
                {isInspectorOpen ? "Hide Inspector" : "Show Inspector"}
              </button>
            )}
          </CommandMenu>
          <CommandMenu label="Help" open={openMenu === "Help"} onOpenChange={(open) => setOpenMenu((current) => open ? "Help" : current === "Help" ? null : current)}>
            <button type="button" onClick={() => setActiveView("help")}>Quick start and Help</button>
            <button type="button" onClick={resetSheet}>Load Demo</button>
          </CommandMenu>
        </nav>
      </header>

      <div className="everydayToolbar">
        <div className="configBar" aria-label="Local configurations">
          <select
            ref={configSelectRef}
            aria-label="Load configuration"
            value={activeConfigId}
            onChange={(event) => handleConfigurationChange(event.target.value)}
          >
            {configurations.map((configuration) => (
              <option key={configuration.id} value={configuration.id}>{configuration.name}</option>
            ))}
          </select>
          <input
            aria-label="Configuration name"
            value={configName}
            onChange={(event) => {
              setConfigName(event.target.value);
              setIsDirty(true);
            }}
          />
          {isDirty && <span>Unsaved</span>}
        </div>
        <div className="toolbarActions">
          <button type="button" onClick={saveConfiguration}>Save</button>
          <button type="button" onClick={undoCells} disabled={undoStack.length === 0}>Undo</button>
          <button type="button" onClick={redoCells} disabled={redoStack.length === 0}>Redo</button>
        </div>
        <div className="viewTabs" role="tablist" aria-label="Quoin surfaces">
          <button type="button" data-active={activeView === "sheet"} onClick={() => setActiveView("sheet")}>Sheet</button>
          <button type="button" data-active={activeView === "runner"} onClick={() => setActiveView("runner")}>Runner</button>
        </div>
        <div className="status" data-valid={workbookResult.valid}>
          <strong>{workbookResult.valid ? "Workbook Ready" : "Workbook Error"}</strong>
          {!workbookResult.valid && firstStatusIssue && (
            <button type="button" className="statusIssue" onClick={openWorkbookIssue}>
              {activeSheetHasIssues ? "Active Sheet" : firstStatusIssue.sheetName ?? "Workbook"}
              {` · ${firstStatusIssue.address}: ${firstStatusIssue.message}`}
            </button>
          )}
        </div>
      </div>

      <div className="fileInputHost">
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            className="fileInput"
            onChange={handleImportFileChange}
          />
      </div>

      {(pendingImport || importError) && (
        <section className="importPanel" aria-label="Workbook import">
          <div className="importPanelHeader">
            <div>
              <p className="eyebrow">Workbook Import</p>
              <h2>{pendingImport ? pendingImport.fileName : "Import Status"}</h2>
            </div>
            {pendingImport && (
              <button type="button" onClick={cancelPendingImport}>
                Cancel
              </button>
            )}
          </div>

          {importError && <p className="importError">{importError}</p>}
          {importMessage && <p className="importMessage">{importMessage}</p>}

          {pendingImport && selectedImportSheet && (
            <div className="importControls">
              <label>
                Open Sheet
                <select value={selectedImportSheetId} onChange={(event) => setSelectedImportSheetId(event.target.value)}>
                  {pendingImport.sheets.map((sheet) => (
                    <option key={sheet.id} value={sheet.id}>
                      {sheet.name}
                    </option>
                  ))}
                </select>
              </label>

              <div className="importStats">
                <span>{selectedImportSheet.cells.length} cells</span>
                <span>{selectedImportSheet.cells.filter((cell) => cell.kind === "formula").length} formulas</span>
                <span>{pendingImport.names.filter((name) => !name.sheetName || name.sheetName === selectedImportSheet.name).length} names</span>
                <span>{importReviewItems.length} review items</span>
              </div>

              <button type="button" onClick={confirmImportSheet}>
                Import Workbook
              </button>
            </div>
          )}

        </section>
      )}

      {activeView === "sheet" ? (
        <>
          <div className="formulaBar">
            <span>{activeSheet ? `${activeSheet.name}!${selectedAddress}` : selectedAddress}</span>
            <input
              aria-label="Formula bar"
              ref={formulaInputRef}
              value={editingAddress === selectedAddress ? draftEntry : selectedCell.entry}
              onChange={(event) => {
                if (editingAddress !== selectedAddress) setEditingAddress(selectedAddress);
                setDraftEntry(event.target.value);
              }}
              onKeyDown={handleFormulaBarKeyDown}
              onMouseDown={() => {
                setIsFormulaBarActive(true);
              }}
              onBlur={() => {
                if (editingAddress === selectedAddress) commitEditing();
                else setIsFormulaBarActive(false);
              }}
              onFocus={() => {
                setIsFormulaBarActive(true);
                setEditingAddress(selectedAddress);
                setDraftEntry(editingAddress === selectedAddress ? draftEntry : selectedCell.entry);
              }}
              placeholder="Value or formula"
            />
            {editingAddress && draftEntry.trim().startsWith("=") && (
              <div className="referencePopup">
                <div className="referencePopupHeader">Insert reference</div>
                {visibleReferenceOptions
                  .map((option, index) => (
                    <button
                      key={option.key}
                      type="button"
                      data-active={index === activeReferenceIndex}
                      onMouseDown={(event) => {
                        event.preventDefault();
                        insertReference(option.reference);
                      }}
                    >
                      <span>{option.reference}</span>
                      <small>{option.address}{option.value !== null && option.value !== "" ? ` = ${formatCellValue(option.value)}` : ""}</small>
                    </button>
                  ))}
                {visibleReferenceOptions.length === 0 && <p>No matching references</p>}
              </div>
            )}
          </div>

          <SheetStrip
            activeSheetId={activeSheetId}
            addSheet={addSheet}
            renameSheet={renameSheet}
            sheets={visibleSheets}
            switchSheet={switchSheet}
          />

          <div className="authoringLayout" data-inspector-open={isInspectorOpen}>
            <section
              className="spreadsheetFrame"
              aria-label="Quoin spreadsheet grid"
              onScroll={(event) => {
                if (rowCount <= 200) return;
                const next = Math.max(0, Math.min(rowCount - 100, Math.floor(event.currentTarget.scrollTop / 26) - 20));
                setGridWindowStart(next);
              }}
            >
              {!isInspectorOpen && (
                <button className="inspectorReopen" type="button" onClick={() => setIsInspectorOpen(true)}>
                  Show Inspector
                </button>
              )}
              <div
                className="spreadsheetGrid"
                style={{
                  gridTemplateColumns: `36px ${columnWidths.map((width) => `${width}px`).join(" ")}`,
                  minWidth: 36 + columnWidths.reduce((total, width) => total + width, 0),
                }}
              >
                <div className="sheetCorner" />
                {columns.map((column) => (
                  <div className="columnHeader" key={column}>{column}</div>
                ))}

                {rowCount > 200 && gridWindowStart > 0 && (
                  <div className="gridRowSpacer" style={{ height: gridWindowStart * 26 }} />
                )}
                {Array.from(
                  { length: rowCount > 200 ? Math.min(100, rowCount - gridWindowStart) : rowCount },
                  (_, rowIndex) => {
                  const rowNumber = (rowCount > 200 ? gridWindowStart : 0) + rowIndex + 1;
                  return (
                    <Row
                      columns={columns}
                      cells={cells}
                      cellRefs={cellRefs}
                      commitEditing={commitEditing}
                      displayValues={displayValues}
                      draftEntry={draftEntry}
                      editInputRef={editInputRef}
                      editingAddress={editingAddress}
                      handleCellClick={handleCellClick}
                      handleCellMouseDown={handleCellMouseDown}
                      handleGridKeyDown={handleGridKeyDown}
                      handleGridPaste={handleGridPaste}
                      issueMap={issueMap}
                      key={rowNumber}
                      rowNumber={rowNumber}
                      selectedAddress={selectedAddress}
                      setDraftEntry={setDraftEntry}
                      startEditing={startEditing}
                      updateCell={updateCell}
                    />
                  );
                })}
                {rowCount > 200 && gridWindowStart + 100 < rowCount && (
                  <div className="gridRowSpacer" style={{ height: (rowCount - gridWindowStart - 100) * 26 }} />
                )}
              </div>
            </section>

            {isInspectorOpen && (
              <Inspector
                clearCell={clearCell}
                closeInspector={() => setIsInspectorOpen(false)}
                dependencySummary={dependencySummary}
                displayValue={displayValues[selectedAddress] ?? ""}
                contextKey={`${activeSheetId}:${selectedAddress}`}
                selectedAddress={selectedAddress}
                selectedCell={selectedCell}
                selectedIssues={selectedIssues}
                updateCell={updateCell}
                updateLookup={updateLookup}
                visibilityControls={visibilityControls}
              />
            )}
          </div>
        </>
      ) : activeView === "runner" ? (
        <RunnerPreview
          backToSheet={() => setActiveView("sheet")}
          configName={configName}
          hasRunnerOverrides={hasRunnerOverrides}
          resetRunnerPreview={resetRunnerPreview}
          runnerSheets={runnerSheets}
          updateRunnerCell={updateRunnerCell}
        />
      ) : (
        <HelpPanel loadDemo={resetSheet} />
      )}

    </>
  );
}

function DraftInput({ contextKey, onCommit, value, ...props }: {
  contextKey: string;
  onCommit: (value: string) => void;
  value: string;
} & Omit<ComponentProps<"input">, "onChange" | "onBlur" | "value">) {
  const [draft, setDraft] = useState(value);
  const draftRef = useRef(draft);
  const initialRef = useRef(value);
  const commitRef = useRef(onCommit);
  draftRef.current = draft;
  commitRef.current = onCommit;

  useEffect(() => {
    setDraft(value);
    draftRef.current = value;
    initialRef.current = value;
    return () => {
      if (draftRef.current !== initialRef.current) commitRef.current(draftRef.current);
    };
  }, [contextKey, value]);

  function commit() {
    if (draftRef.current === initialRef.current) return;
    commitRef.current(draftRef.current);
    initialRef.current = draftRef.current;
  }

  return <input {...props} value={draft} onChange={(event) => setDraft(event.target.value)} onBlur={commit} />;
}

function DraftTextarea({ contextKey, onCommit, value, ...props }: {
  contextKey: string;
  onCommit: (value: string) => void;
  value: string;
} & Omit<ComponentProps<"textarea">, "onChange" | "onBlur" | "value">) {
  const [draft, setDraft] = useState(value);
  const draftRef = useRef(draft);
  const initialRef = useRef(value);
  const commitRef = useRef(onCommit);
  draftRef.current = draft;
  commitRef.current = onCommit;

  useEffect(() => {
    setDraft(value);
    draftRef.current = value;
    initialRef.current = value;
    return () => {
      if (draftRef.current !== initialRef.current) commitRef.current(draftRef.current);
    };
  }, [contextKey, value]);

  function commit() {
    if (draftRef.current === initialRef.current) return;
    commitRef.current(draftRef.current);
    initialRef.current = draftRef.current;
  }

  return <textarea {...props} value={draft} onChange={(event) => setDraft(event.target.value)} onBlur={commit} />;
}

function Inspector({
  clearCell,
  closeInspector,
  dependencySummary,
  displayValue,
  contextKey,
  selectedAddress,
  selectedCell,
  selectedIssues,
  updateCell,
  updateLookup,
  visibilityControls,
}: {
  clearCell: (address: string) => void;
  closeInspector: () => void;
  dependencySummary: DependencySummary;
  displayValue: CellValue;
  contextKey: string;
  selectedAddress: string;
  selectedCell: GridCell;
  selectedIssues: string[];
  updateCell: (address: string, patch: Partial<GridCell>) => void;
  updateLookup: (patch: Partial<LookupConfig>) => void;
  visibilityControls: VisibilityControlOption[];
}) {
  const selectedVisibilityControl = selectedCell.visibilityCondition?.source
    ? visibilityControls.find((control) => control.name === selectedCell.visibilityCondition?.source)
    : null;
  const hasMissingVisibilitySource = Boolean(selectedCell.visibilityCondition?.source && !selectedVisibilityControl);

  return (
    <aside className="inspector" aria-label="Selected cell inspector">
      <div className="inspectorHeader">
        <div>
          <span>{selectedAddress}</span>
          <h2>{selectedCell.name || "Normal cell"}</h2>
        </div>
        <div className="inspectorActions">
          {selectedCell.name ? (
            <span className="smartPill">{selectedCell.role}</span>
          ) : (
            <span className="normalPill">Normal</span>
          )}
          <button type="button" onClick={closeInspector}>Hide Inspector</button>
          <button type="button" onClick={() => clearCell(selectedAddress)}>Clear Cell</button>
        </div>
      </div>

      {!selectedCell.name && (
        <div className="activationBox">
          <strong>Normal spreadsheet cell</strong>
          <p>Start by typing values and formulas in the grid. Add a Smart Cell name only when this cell needs runner visibility, metadata, or named formula references.</p>
        </div>
      )}

      <div className="inspectorSection">
        <div className="sectionTitle">
          <strong>{selectedAddress} · {selectedCell.entry.trim().startsWith("=") ? "Formula" : prettifyName(selectedCell.type)}</strong>
          <span>{selectedCell.name ? "Smart Cell identity" : "Name this cell when it needs formula references or Runner behavior."}</span>
        </div>

        <div className="detectedTypeRow">
          <span>{selectedCell.entry.trim().startsWith("=") ? "Formula" : "Detected type"}</span>
          <strong>{selectedCell.entry.trim().startsWith("=") ? "formula" : selectedCell.type}</strong>
        </div>

        <label>
          Smart Cell Name
          <DraftInput
            contextKey={`${contextKey}:name`}
            placeholder="example: wall_height"
            value={selectedCell.name}
            onCommit={(value) => {
              const name = sanitizeName(value);
              updateCell(selectedAddress, {
                name,
                surfaced: name ? selectedCell.surfaced : false,
              });
            }}
          />
        </label>

      </div>

      {selectedCell.name && (
        <details className="inspectorDisclosure smartSection" open>
          <summary>
            <span>Smart Cell</span>
            <small>Role and Runner surfacing</small>
          </summary>
          <div className="inspectorDisclosureBody">
          <div className="inspectorGrid">
            <label>
              Role
              <select
                value={selectedCell.role}
                onChange={(event) => {
                  const role = event.target.value as SmartCellRole;
                  updateCell(selectedAddress, { role, lookup: role === "lookup" || role === "action" ? selectedCell.lookup ?? starterLookup : selectedCell.lookup });
                }}
              >
                <optgroup label="Core">
                  {coreRoleOptions.map((role) => (
                    <option key={role} value={role}>{role}</option>
                  ))}
                </optgroup>
                <optgroup label="Advanced prototype">
                  {advancedRoleOptions.map((role) => (
                    <option key={role} value={role}>{role}</option>
                  ))}
                </optgroup>
              </select>
            </label>
            <label>
              Value Type
              <select
                value={selectedCell.type}
                onChange={(event) => updateCell(selectedAddress, { type: event.target.value as SmartCellType })}
              >
                {typeOptions.map((type) => (
                  <option key={type} value={type}>{type}</option>
              ))}
              </select>
            </label>
          </div>

          {selectedCell.role === "input" && (
            <details className="nestedDisclosure" open>
              <summary>Input settings</summary>
              <div className="nestedDisclosureBody">
                <label>
                  Input Control
                  <select
                    value={selectedCell.inputControl}
                    onChange={(event) => {
                      const inputControl = event.target.value as InputControl;
                      const options = inputControl === "dropdown" && selectedCell.inputOptions.length === 0 && selectedCell.entry.trim()
                        ? [selectedCell.entry.trim()]
                        : selectedCell.inputOptions;
                      updateCell(selectedAddress, {
                        inputControl,
                        inputOptions: inputControl === "dropdown" ? options : [],
                        type: inputControl === "checkbox" ? "boolean" : selectedCell.type,
                      });
                    }}
                  >
                    {inputControlOptions.map((inputControl) => (
                      <option key={inputControl} value={inputControl}>{labelForInputControl(inputControl)}</option>
                    ))}
                  </select>
                </label>

                {selectedCell.inputControl === "dropdown" && (
                  <label>
                    Dropdown Options
                    <DraftTextarea
                      contextKey={`${contextKey}:options`}
                      value={selectedCell.inputOptions.join("\n")}
                      onCommit={(value) => updateCell(selectedAddress, { inputOptions: splitInputOptions(value) })}
                      placeholder="One short option per line, or comma-separated. Leave blank for free text."
                      rows={3}
                    />
                    <span>Use embedded options for short lists. Longer lists should come from visible reference data in a later reference-table workflow.</span>
                  </label>
                )}
              </div>
            </details>
          )}

          <div className="surfaceRow">
            <label className="checkLabel">
              <input
                checked={selectedCell.surfaced}
                type="checkbox"
                onChange={(event) => updateCell(selectedAddress, { surfaced: event.target.checked })}
              />
              Surface to runner
            </label>
            <span>{selectedCell.surfaced ? "Visible in Runner Preview" : "Authoring only"}</span>
          </div>

          {selectedCell.surfaced && (
            <details className="nestedDisclosure">
              <summary>Runner appearance</summary>
              <div className="nestedDisclosureBody">
              <label>
                Display Label
                <DraftInput
                  contextKey={`${contextKey}:label`}
                  placeholder={prettifyName(selectedCell.name)}
                  value={selectedCell.label}
                  onCommit={(value) => updateCell(selectedAddress, { label: value })}
                />
              </label>
              <label>
                Runner Section
                <DraftInput
                  contextKey={`${contextKey}:section`}
                  value={selectedCell.runnerSection}
                  onCommit={(value) => updateCell(selectedAddress, { runnerSection: value })}
                  placeholder="Example: Dormers"
                />
              </label>
              <label>
                Conditional
                <select
                  value={selectedCell.visibilityCondition?.source ?? ""}
                  onChange={(event) => {
                    const source = event.target.value;
                    updateCell(selectedAddress, {
                      visibilityCondition: source
                        ? { source, operator: "equals", value: true }
                        : undefined,
                    });
                  }}
                >
                  <option value="">Always show</option>
                  {visibilityControls
                    .filter((control) => control.name !== selectedCell.name)
                    .map((control) => (
                      <option key={control.name} value={control.name}>
                        Show when {control.label} is checked
                      </option>
                    ))}
                </select>
                <span>
                  {selectedCell.visibilityCondition?.source
                    ? selectedVisibilityControl
                      ? `Controlled by ${selectedVisibilityControl.sheetName} / ${selectedVisibilityControl.label}.`
                      : "The selected control no longer exists or is not a boolean input."
                    : "Use a surfaced checkbox input to show optional runner fields only when needed."}
                </span>
              </label>
              {hasMissingVisibilitySource && (
                <div className="issueBox">
                  <p>This Smart Cell is hidden in Runner Preview until its missing condition source is repaired.</p>
                </div>
              )}
              </div>
            </details>
          )}

          <details className="nestedDisclosure">
            <summary>Notes &amp; advanced</summary>
            <div className="nestedDisclosureBody">
              <label>
                Internal Annotation
                <DraftTextarea
                  contextKey={`${contextKey}:annotation`}
                  value={selectedCell.annotation}
                  onCommit={(value) => updateCell(selectedAddress, { annotation: value })}
                  placeholder="Internal note about what this cell means"
                  rows={4}
                />
              </label>

          {(selectedCell.role === "validation" || selectedCell.role === "compliance") && (
            <label>
              Runner Message
              <DraftTextarea
                contextKey={`${contextKey}:ruleMessage`}
                value={selectedCell.ruleMessage}
                onCommit={(value) => updateCell(selectedAddress, { ruleMessage: value })}
                placeholder={selectedCell.role === "validation" ? "Example: Input is outside the approved range." : "Example: Manual review is recommended."}
                rows={3}
              />
            </label>
          )}
            </div>
          </details>
          </div>
        </details>
      )}

      {selectedCell.role === "lookup" && selectedCell.name && (
        <LookupEditor lookup={selectedCell.lookup ?? starterLookup} updateLookup={updateLookup} />
      )}

      {selectedCell.role === "compliance" && selectedCell.name && (
        <div className="inspectorNote">
          <strong>Compliance Rule</strong>
          <p>Warn when this condition is true. The run can continue, and the runner sees the Runner Message.</p>
        </div>
      )}

      {selectedCell.role === "validation" && selectedCell.name && (
        <div className="inspectorNote">
          <strong>Validation Rule</strong>
          <p>Fail the run when this condition is false. The math still runs, and the runner sees the Runner Message.</p>
        </div>
      )}

      <div className="inspectorNote">
        <strong>Current Value</strong>
        <p>{formatCellValue(displayValue)}</p>
      </div>

      <details className="inspectorDisclosure">
        <summary>
          <span>Dependencies</span>
          <small>Upstream and downstream references</small>
        </summary>
        <div className="dependencyPanel">
        <div>
          <strong>Depends On</strong>
          {dependencySummary.dependencies.length === 0 ? (
            <p>No upstream references.</p>
          ) : (
            dependencySummary.dependencies.map((item) => (
              <p key={`${item.address}-${item.reference}`}>
                <span>{item.reference}</span>
                {item.label}
              </p>
            ))
          )}
        </div>
        <div>
          <strong>Used By</strong>
          {dependencySummary.dependents.length === 0 ? (
            <p>No downstream cells.</p>
          ) : (
            dependencySummary.dependents.map((item) => (
              <p key={`${item.address}-${item.reference}`}>
                <span>{item.address}</span>
                {item.label}
              </p>
            ))
          )}
        </div>
        </div>
      </details>

      {selectedIssues.length > 0 && (
        <div className="issueBox">
          <strong>Cell Issue</strong>
          {selectedIssues.map((issue) => (
            <p key={issue}>{issue}</p>
          ))}
        </div>
      )}
    </aside>
  );
}

function RunnerPreview({
  backToSheet,
  configName,
  hasRunnerOverrides,
  resetRunnerPreview,
  updateRunnerCell,
  runnerSheets,
}: {
  backToSheet: () => void;
  configName: string;
  hasRunnerOverrides: boolean;
  resetRunnerPreview: () => void;
  runnerSheets: RunnerSheetContext[];
  updateRunnerCell: (sheetId: string, address: string, entry: string) => void;
}) {
  const showSheetGroups = runnerSheets.filter((sheet) => sheet.surfacedCells.length > 0).length > 1;
  const inputGroups = runnerSheets.map((sheet) => ({
    ...sheet,
    items: sheet.surfacedCells.filter((cell) => cell.role === "input"),
  })).filter((sheet) => sheet.items.length > 0);
  const outputGroups = runnerSheets.map((sheet) => ({
    ...sheet,
    items: sheet.surfacedCells.filter((cell) => cell.role !== "input" && cell.role !== "validation" && cell.role !== "compliance" && cell.role !== "action"),
  })).filter((sheet) => sheet.items.length > 0);
  const actionGroups = runnerSheets.map((sheet) => ({
    ...sheet,
    items: sheet.surfacedCells.filter((cell) => cell.role === "action"),
  })).filter((sheet) => sheet.items.length > 0);
  const warningGroups = runnerSheets.map((sheet) => ({
    ...sheet,
    items: sheet.result.warnings.filter((warning) => sheet.surfacedCells.some((cell) => cell.address === warning.address)),
  })).filter((sheet) => sheet.items.length > 0);
  const validationGroups = runnerSheets.map((sheet) => ({
    ...sheet,
    items: sheet.validationStates,
  })).filter((sheet) => sheet.items.length > 0);
  const visibleRunnerValid = !validationGroups.some((group) => group.items.some((rule) => rule.state === "fail" || rule.state === "error"));

  return (
    <section className="runnerPreview">
      <div className="runnerHeader">
        <div>
          <p className="eyebrow">Runner Preview</p>
          <h2>{configName || "Untitled Configuration"}</h2>
          <p className="runnerSessionNote">Runner edits are temporary and do not rewrite the Sheet defaults.</p>
        </div>
        <div className="runnerHeaderActions">
          <span data-valid={visibleRunnerValid}>{visibleRunnerValid ? "Ready" : "Failed Validation"}</span>
          <button type="button" onClick={backToSheet}>Back to Sheet</button>
          <button type="button" onClick={resetRunnerPreview} disabled={!hasRunnerOverrides}>Reset Preview</button>
        </div>
      </div>

      <div className={`runnerGrid ${outputGroups.length === 0 ? "runnerGridSingle" : ""}`}>
        <div className="runnerPanel">
          <h3>Inputs</h3>
          {inputGroups.length === 0 ? (
            <p className="runnerEmpty">No surfaced inputs.</p>
          ) : (
            inputGroups.map((group) => (
              <RunnerSheetGroup key={group.sheetId} showHeading={showSheetGroups} sheetName={group.sheetName}>
                {groupItemsByRunnerSection(group.items, (cell) => cell).map((section) => (
                  <RunnerSectionGroup key={section.label || "default"} label={section.label}>
                    {section.items.map((cell) => (
                      <label key={`${group.sheetId}-${cell.address}`}>
                        {labelForCell(cell)}
                        {isDropdownCell(cell) ? (
                          <select value={cell.entry} onChange={(event) => updateRunnerCell(group.sheetId, cell.address, event.target.value)}>
                            {!cell.entry && <option value="">Choose...</option>}
                            {dropdownOptionsForCell(cell).map((option) => (
                              <option key={option} value={option}>{prettifyName(option)}</option>
                            ))}
                          </select>
                        ) : isCheckboxCell(cell) ? (
                          <span className="runnerCheckbox">
                            <input
                              checked={parseCellValue(cell.entry, "boolean") === true}
                              type="checkbox"
                              onChange={(event) => updateRunnerCell(group.sheetId, cell.address, event.target.checked ? "true" : "false")}
                            />
                            {parseCellValue(cell.entry, "boolean") === true ? "Checked" : "Unchecked"}
                          </span>
                        ) : (
                          <input value={cell.entry} onChange={(event) => updateRunnerCell(group.sheetId, cell.address, event.target.value)} />
                        )}
                        {cell.annotation && <small>{cell.annotation}</small>}
                      </label>
                    ))}
                  </RunnerSectionGroup>
                ))}
              </RunnerSheetGroup>
            ))
          )}
        </div>

        {outputGroups.length > 0 && (
          <div className="runnerPanel">
            <h3>Outputs</h3>
            {outputGroups.map((group) => (
              <RunnerSheetGroup key={group.sheetId} showHeading={showSheetGroups} sheetName={group.sheetName}>
                {groupItemsByRunnerSection(group.items, (cell) => cell).map((section) => (
                  <RunnerSectionGroup key={section.label || "default"} label={section.label}>
                    {section.items.map((cell) => (
                      <div className="runnerResult" key={`${group.sheetId}-${cell.address}`}>
                        <span>{labelForCell(cell)}</span>
                        <strong>{formatCellValue(group.displayValues[cell.address] ?? null)}</strong>
                        {cell.annotation && <small>{cell.annotation}</small>}
                      </div>
                    ))}
                  </RunnerSectionGroup>
                ))}
              </RunnerSheetGroup>
            ))}
          </div>
        )}
      </div>

      {(actionGroups.length > 0 || warningGroups.length > 0) && (
        <div className="runnerMessages">
          {actionGroups.length > 0 && (
            <div>
              <h3>Shop Actions</h3>
              {actionGroups.map((group) => (
                <RunnerSheetGroup key={group.sheetId} showHeading={showSheetGroups} sheetName={group.sheetName}>
                  {groupItemsByRunnerSection(group.items, (cell) => cell).map((section) => (
                    <RunnerSectionGroup key={section.label || "default"} label={section.label}>
                      {section.items.map((cell) => (
                        <p data-state="action" key={`${group.sheetId}-${cell.address}`}>
                          <strong>ACTION</strong>
                          {formatCellValue(group.displayValues[cell.address] ?? null) || labelForCell(cell)}
                        </p>
                      ))}
                    </RunnerSectionGroup>
                  ))}
                </RunnerSheetGroup>
              ))}
            </div>
          )}
          {warningGroups.length > 0 && (
            <div>
              <h3>Review Flags</h3>
              {warningGroups.map((group) => (
                <RunnerSheetGroup key={group.sheetId} showHeading={showSheetGroups} sheetName={group.sheetName}>
                  {groupItemsByRunnerSection(group.items, (warning) => getCell(group.cells, warning.address)).map((section) => (
                    <RunnerSectionGroup key={section.label || "default"} label={section.label}>
                      {section.items.map((warning) => (
                        <p data-state="warn" key={warning.cellId}>
                          <strong>WARN</strong>
                          {warning.message}
                        </p>
                      ))}
                    </RunnerSectionGroup>
                  ))}
                </RunnerSheetGroup>
              ))}
            </div>
          )}
        </div>
      )}

      {validationGroups.length > 0 && (
        <div className="runnerMessages runnerMessagesSingle">
          <div>
            <h3>Validation</h3>
            {validationGroups.map((group) => (
              <RunnerSheetGroup key={group.sheetId} showHeading={showSheetGroups} sheetName={group.sheetName}>
                {groupItemsByRunnerSection(group.items, (rule) => getCell(group.cells, rule.address)).map((section) => (
                  <RunnerSectionGroup key={section.label || "default"} label={section.label}>
                    {section.items.map((rule) => {
                      const cell = getCell(group.cells, rule.address);
                      return (
                        <p data-state={rule.state} key={`${group.sheetId}-${rule.address}`}>
                          <strong>{formatCellValue(group.displayValues[rule.address] ?? null)}</strong>
                          {cell.ruleMessage || cell.annotation || labelForCell(cell)}
                        </p>
                      );
                    })}
                  </RunnerSectionGroup>
                ))}
              </RunnerSheetGroup>
            ))}
          </div>
        </div>
      )}
    </section>
  );
}

function RunnerSheetGroup({
  children,
  sheetName,
  showHeading,
}: {
  children: ReactNode;
  sheetName: string;
  showHeading: boolean;
}) {
  return (
    <div className="runnerSheetGroup">
      {showHeading && <h4>{sheetName}</h4>}
      {children}
    </div>
  );
}

function RunnerSectionGroup({
  children,
  label,
}: {
  children: ReactNode;
  label: string;
}) {
  return (
    <div className="runnerSectionGroup">
      {label && <h5>{label}</h5>}
      {children}
    </div>
  );
}

function HelpPanel({ loadDemo }: { loadDemo: () => void }) {
  return (
    <section className="helpPanel" aria-label="Quoin help">
      <div className="helpHeader">
        <p className="eyebrow">Help</p>
        <h2>Build a spreadsheet. Turn it into a Runner.</h2>
        <p>Quoin starts with familiar spreadsheet work, then adds structure only where it helps the person using the finished calculator.</p>
      </div>

      <section className="helpStart" aria-labelledby="help-start-title">
        <div>
          <p className="eyebrow">Start Here</p>
          <h3 id="help-start-title">The Quoin workflow</h3>
          <ol className="helpWorkflow">
            <li><strong>Build</strong><span>Enter values and formulas in normal cells.</span></li>
            <li><strong>Verify</strong><span>Make sure the Sheet calculates correctly.</span></li>
            <li><strong>Name</strong><span>Promote important cells into Smart Cells.</span></li>
            <li><strong>Surface</strong><span>Choose what the runner should see.</span></li>
            <li><strong>Preview</strong><span>Test the generated Runner without changing Sheet defaults.</span></li>
          </ol>
        </div>
        <div className="helpConceptExample">
          <strong>One cell, three decisions</strong>
          <dl>
            <div><dt>Name</dt><dd><code>design_span</code> - safe to use in formulas.</dd></div>
            <div><dt>Display label</dt><dd><code>Design Span (ft)</code> - clear wording for the runner.</dd></div>
            <div><dt>Surface</dt><dd>Include the cell in Runner Preview.</dd></div>
          </dl>
          <button type="button" onClick={loadDemo}>Load the example configuration</button>
        </div>
      </section>

      <div className="helpTaskGrid">
        <article>
          <p className="helpStep">1 / Author</p>
          <h3>Build the Sheet</h3>
          <p>Normal cells hold labels, values, booleans, and formulas. They remain in the authoring grid unless you promote them.</p>
          <ul>
            <li>Type in a cell, double-click it, press Enter, or use the formula bar.</li>
            <li>Use Sheet tabs to organize a workbook and paste tabular data into the grid.</li>
            <li>Use Undo/Redo, Copy/Paste, and Fill Down from the toolbar, menus, or keyboard.</li>
          </ul>
          <details>
            <summary>Keyboard shortcuts</summary>
            <ul>
              <li>Arrow keys move the selection; Enter or F2 starts editing.</li>
              <li>Tab moves right; Shift+Tab moves left.</li>
              <li>Ctrl+Z and Ctrl+Y undo and redo.</li>
              <li>Ctrl+C, Ctrl+V, and Ctrl+D copy, paste, and fill down.</li>
              <li>Delete or Backspace clears the selected cell.</li>
            </ul>
          </details>
        </article>

        <article>
          <p className="helpStep">2 / Structure</p>
          <h3>Create Smart Cells</h3>
          <p>Giving a cell a Smart Cell Name adds formula-safe identity and opens its workflow settings in the inspector.</p>
          <ul>
            <li><strong>Input</strong> accepts an admin or runner value.</li>
            <li><strong>Formula</strong> holds internal calculated logic.</li>
            <li><strong>Output</strong> presents a calculated or entered result.</li>
          </ul>
          <p>Keep names formula-safe; use Display Label for human-facing wording.</p>
          <details>
            <summary>Dropdowns and checkboxes</summary>
            <ul>
              <li>Use Dropdown for a short embedded list of choices.</li>
              <li>Use Checkbox for true/false inputs.</li>
              <li>Both controls appear in the Sheet and Runner Preview.</li>
              <li>Long option lists should eventually come from visible reference data.</li>
            </ul>
          </details>
        </article>

        <article>
          <p className="helpStep">3 / Preview</p>
          <h3>Build the Runner</h3>
          <p>Runner Preview is generated from surfaced Smart Cells across the workbook. Normal and unsurfaced cells remain authoring-only.</p>
          <ul>
            <li>Runner edits are temporary and do not rewrite Sheet defaults.</li>
            <li>Reset Preview clears temporary values.</li>
            <li>Outputs update from runner-session inputs.</li>
            <li>Back to Sheet returns to authoring.</li>
          </ul>
          <details>
            <summary>Conditional sections</summary>
            <ul>
              <li>A surfaced checkbox can control whether other surfaced fields appear.</li>
              <li>Runner Section gives related fields a shared heading.</li>
              <li>Hidden Runner inputs return to their Sheet defaults.</li>
              <li>Visibility does not change workbook math; use <code>IF</code> when optional inputs should change totals.</li>
            </ul>
          </details>
        </article>

        <article>
          <p className="helpStep">4 / Bring Work In</p>
          <h3>Import a Workbook</h3>
          <p>Import Workbook creates a new browser-local configuration from an <code>.xlsx</code> file. Preserve the calculator first, then structure it with Quoin.</p>
          <ol>
            <li>Import the workbook and choose the first Sheet.</li>
            <li>Review preserved Sheets, values, formulas, names, and dropdowns.</li>
            <li>Inspect review items and repair formulas where needed.</li>
            <li>Promote important cells into Smart Cells.</li>
          </ol>
          <details>
            <summary>What import preserves and reviews</summary>
            <ul>
              <li>Sheets, ordinary values, formulas, and supported dropdown lists are preserved.</li>
              <li>Safe single-cell workbook names can become Smart Cell names.</li>
              <li>Merged ranges, external links, structured references, spill markers, and unsupported formulas produce review items.</li>
              <li>Unsupported formulas remain visible instead of being silently discarded.</li>
            </ul>
          </details>
        </article>
      </div>

      <section className="helpReference">
        <div className="helpSectionHeading">
          <p className="eyebrow">Reference</p>
          <h3>Formula Reference</h3>
          <p>Quoin supports common deterministic spreadsheet formulas. Formulas begin with <code>=</code>.</p>
        </div>
        <details className="helpDisclosure" open>
          <summary>References and ranges</summary>
          <div><ul>
            <li><code>=A1+B1</code> uses coordinates on the active Sheet.</li>
            <li><code>=design_span * design_plf</code> uses workbook-scoped Smart Cell names.</li>
            <li><code>=Inputs!B2 * Inputs!B3</code> uses another Sheet.</li>
            <li><code>='Input Data'!B2 * 3</code> quotes a Sheet name containing spaces.</li>
            <li><code>=SUM(A1:B3)</code> and <code>=COUNT(Loads!B3:B5)</code> use ranges.</li>
          </ul></div>
        </details>
        <details className="helpDisclosure">
          <summary>Common functions</summary>
          <div className="helpTableScroll"><table className="helpFunctionTable">
            <thead><tr><th>Group</th><th>Supported examples</th><th>Notes</th></tr></thead>
            <tbody>
              <tr><td>Aggregates</td><td><code>SUM</code>, <code>COUNT</code>, <code>AVERAGE</code>, <code>MIN</code>, <code>MAX</code></td><td><code>COUNT</code> ignores blanks and text.</td></tr>
              <tr><td>Numeric</td><td><code>ROUND</code>, <code>ROUNDUP</code>, <code>ABS</code>, <code>SQRT</code>, <code>CEIL</code>, <code>FLOOR</code></td><td>Common numeric cleanup.</td></tr>
              <tr><td>Conditions</td><td><code>IF</code> and comparisons</td><td>Numeric and text comparisons.</td></tr>
              <tr><td>Lookups</td><td>Exact <code>VLOOKUP</code>, exact/default <code>XLOOKUP</code></td><td>Approximate matching requires review.</td></tr>
            </tbody>
          </table></div>
        </details>
        <details className="helpDisclosure">
          <summary>Conditions and IF</summary>
          <div><ul>
            <li><code>=design_span&gt;14</code> returns true or false.</li>
            <li><code>=IF(A1&gt;10, "review", "ok")</code> chooses between two results.</li>
            <li><code>=IF(material="premium", 6, 3)</code> compares text.</li>
            <li><code>=IF(material&lt;&gt;"premium", "standard", "review")</code> tests inequality.</li>
          </ul></div>
        </details>
        <details className="helpDisclosure">
          <summary>VLOOKUP and XLOOKUP</summary>
          <div>
            <p>Lookup formulas work in normal cells; a lookup Smart Cell is not required.</p>
            <ul>
              <li><code>=VLOOKUP(A1,Data!A1:B200,2,FALSE)</code> performs an exact lookup.</li>
              <li><code>=VLOOKUP(A1&amp;"|"&amp;B1,'Beam Data'!F2:G9999,2,FALSE)</code> supports helper keys.</li>
              <li><code>=XLOOKUP(A1,Data!A1:A200,Data!B1:B200)</code> uses default exact matching.</li>
              <li>Approximate and non-exact modes are preserved for review.</li>
            </ul>
          </div>
        </details>
        <details className="helpDisclosure">
          <summary>Editing, errors, and imported formula review</summary>
          <div><ul>
            <li>The formula popup suggests Smart Cell names and populated coordinates.</li>
            <li>Copy and Fill Down adjust coordinates while leaving Smart Cell names unchanged.</li>
            <li>Renaming a Sheet updates direct cross-Sheet references.</li>
            <li>Deleted references become <code>#REF!</code> and produce a visible engine error.</li>
            <li>An invalid formula does not stop unrelated formulas from calculating.</li>
            <li><code>IFERROR</code>, <code>IFNA</code>, <code>INDIRECT</code>, <code>OFFSET</code>, criteria aggregates, and date functions currently require review or an explicit rewrite.</li>
          </ul></div>
        </details>
      </section>

      <section className="helpCurrentState">
        <div className="helpSectionHeading">
          <p className="eyebrow">Storage &amp; Scope</p>
          <h3>Saving and Current Limits</h3>
        </div>
        <div className="helpStatusGrid">
          <article>
            <h4>Available now</h4>
            <ul>
              <li>Browser-local configurations with New, Save, Duplicate, Delete, and rename.</li>
              <li>Multiple Sheets, workbook import, Smart Cells, and Runner Preview.</li>
              <li>Exact lookup formulas and embedded short dropdown lists.</li>
            </ul>
          </article>
          <article>
            <h4>Not available yet</h4>
            <ul>
              <li>Database persistence, accounts, or permissions.</li>
              <li>Published versions, saved executions, or audit reports.</li>
              <li>First-class reference tables and live CSV-backed options.</li>
              <li>External action automation.</li>
            </ul>
          </article>
        </div>
        <details className="helpDisclosure">
          <summary>Advanced prototype roles</summary>
          <div><p>Lookup, action, validation, and compliance remain available for prototype work. The main workflow should use inputs, formulas, outputs, surfaced text, and Conditional Runner Sections until those roles are redesigned around reference data, automation, and formal run status.</p></div>
        </details>
      </section>

      <section className="helpFooter">
        <div>
          <h3>Want to see the full workflow?</h3>
          <p>Load the example, inspect its formulas and Smart Cells, then open Runner Preview.</p>
        </div>
        <button type="button" onClick={loadDemo}>Load Demo</button>
      </section>
    </section>
  );
}

function Row({
  columns,
  cells,
  cellRefs,
  commitEditing,
  displayValues,
  draftEntry,
  editInputRef,
  editingAddress,
  handleCellClick,
  handleCellMouseDown,
  handleGridKeyDown,
  handleGridPaste,
  issueMap,
  rowNumber,
  selectedAddress,
  setDraftEntry,
  startEditing,
  updateCell,
}: {
  columns: string[];
  cells: Record<string, GridCell>;
  cellRefs: React.MutableRefObject<Record<string, HTMLDivElement | null>>;
  commitEditing: (nextAddress?: string) => void;
  displayValues: Record<string, CellValue>;
  draftEntry: string;
  editInputRef: React.RefObject<HTMLInputElement | null>;
  editingAddress: string | null;
  handleCellClick: (address: string) => void;
  handleCellMouseDown: (event: React.MouseEvent<HTMLDivElement>, address: string) => void;
  handleGridKeyDown: (event: React.KeyboardEvent<HTMLDivElement>, address: string) => void;
  handleGridPaste: (event: React.ClipboardEvent<HTMLDivElement>, address: string) => void;
  issueMap: Map<string, string[]>;
  rowNumber: number;
  selectedAddress: string;
  setDraftEntry: (value: string) => void;
  startEditing: (address: string, replacement?: string) => void;
  updateCell: (address: string, patch: Partial<GridCell>) => void;
}) {
  return (
    <>
      <div className="rowHeader">{rowNumber}</div>
      {columns.map((column) => {
        const address = `${column}${rowNumber}`;
        const cell = getCell(cells, address);
        const selected = selectedAddress === address;
        const editing = editingAddress === address;
        const issues = issueMap.get(address) ?? [];
        const hasDropdown = isDropdownCell(cell);
        const hasCheckbox = isCheckboxCell(cell);
        const dropdownOptions = dropdownOptionsForCell(cell);
        return (
          <div
            className="gridCell"
            data-dropdown={hasDropdown}
            data-checkbox={hasCheckbox}
            data-editing={editing}
            data-issue={issues.length > 0}
            data-role={cell.name ? cell.role : "normal"}
            data-selected={selected}
            data-smart={Boolean(cell.name)}
            key={address}
            onClick={() => handleCellClick(address)}
            onDoubleClick={() => startEditing(address)}
            onKeyDown={(event) => handleGridKeyDown(event, address)}
            onMouseDown={(event) => handleCellMouseDown(event, address)}
            onPaste={(event) => handleGridPaste(event, address)}
            ref={(node) => {
              cellRefs.current[address] = node;
            }}
            tabIndex={0}
          >
            {editing ? (
              <input
                aria-label={address}
                ref={editInputRef}
                value={draftEntry}
                onBlur={() => editingAddress === address && commitEditing()}
                onChange={(event) => setDraftEntry(event.target.value)}
              />
            ) : hasDropdown ? (
              <select
                aria-label={`${address} dropdown`}
                className="gridDropdown"
                value={cell.entry}
                onClick={(event) => event.stopPropagation()}
                onFocus={() => handleCellClick(address)}
                onKeyDown={(event) => event.stopPropagation()}
                onChange={(event) => updateCell(address, { entry: event.target.value })}
              >
                {!cell.entry && <option value="">Choose...</option>}
                {dropdownOptions.map((option) => (
                  <option key={option} value={option}>{prettifyName(option)}</option>
                ))}
              </select>
            ) : hasCheckbox ? (
              <label className="gridCheckbox" onClick={(event) => event.stopPropagation()}>
                <input
                  aria-label={`${address} checkbox`}
                  checked={parseCellValue(cell.entry, "boolean") === true}
                  type="checkbox"
                  onFocus={() => handleCellClick(address)}
                  onKeyDown={(event) => event.stopPropagation()}
                  onChange={(event) => updateCell(address, { entry: event.target.checked ? "true" : "false" })}
                />
              </label>
            ) : (
              <span className="cellDisplay">{formatCellValue(displayValues[address] ?? null)}</span>
            )}
            {cell.name && <span className="smartMarker">{cell.name}</span>}
          </div>
        );
      })}
    </>
  );
}

function LookupEditor({ lookup, updateLookup }: { lookup: LookupConfig; updateLookup: (patch: Partial<LookupConfig>) => void }) {
  const inputs = normalizeLookupInputs(lookup);
  const tableColumns = [...inputs.map((input) => input.column), lookup.outputColumn];

  function updateRow(index: number, column: string, value: string) {
    const rows = lookup.rows.map((row, rowIndex) => {
      if (rowIndex !== index) return row;
      return { ...row, [column]: parseLooseValue(value) };
    });
    updateLookup({ rows });
  }

  function updateInput(index: number, patch: Partial<{ column: string; reference: string }>) {
    const nextInputs = inputs.map((input, inputIndex) => {
      if (inputIndex !== index) return input;
      return { ...input, ...patch };
    });
    updateLookup({
      inputColumn: nextInputs[0]?.column ?? "",
      inputReference: nextInputs[0]?.reference ?? "",
      inputs: nextInputs,
    });
  }

  function addInput() {
    updateLookup({ inputs: [...inputs, { column: "condition", reference: "" }] });
  }

  function removeInput(index: number) {
    const nextInputs = inputs.filter((_, inputIndex) => inputIndex !== index);
    updateLookup({
      inputColumn: nextInputs[0]?.column ?? "",
      inputReference: nextInputs[0]?.reference ?? "",
      inputs: nextInputs,
    });
  }

  function addRow() {
    updateLookup({ rows: [...lookup.rows, Object.fromEntries(tableColumns.map((column) => [column, ""]))] });
  }

  function handlePaste(event: React.ClipboardEvent<HTMLInputElement>, startRow: number, startColumn: string) {
    const text = event.clipboardData.getData("text/plain");
    if (!text.includes("\t") && !text.includes("\n")) return;

    event.preventDefault();
    const pastedRows = parsePastedRows(text);
    if (pastedRows.length === 0) return;

    const rowsToApply = firstRowLooksLikeLookupHeader(pastedRows[0], tableColumns) ? pastedRows.slice(1) : pastedRows;
    if (rowsToApply.length === 0) return;

    const startColumnIndex = tableColumns.indexOf(startColumn);
    const nextRows = lookup.rows.map((row) => ({ ...row }));

    rowsToApply.forEach((pastedRow, rowOffset) => {
      const rowIndex = startRow + rowOffset;
      while (nextRows.length <= rowIndex) {
        nextRows.push(Object.fromEntries(tableColumns.map((column) => [column, ""])));
      }

      pastedRow.forEach((value, columnOffset) => {
        const column = tableColumns[startColumnIndex + columnOffset];
        if (!column) return;
        nextRows[rowIndex] = { ...nextRows[rowIndex], [column]: parseLooseValue(value) };
      });
    });

    updateLookup({ rows: nextRows });
  }

  return (
    <div className="lookupEditor">
      <div className="lookupHeader">
        <strong>Lookup Table</strong>
        <button type="button" onClick={addRow}>Add Row</button>
      </div>

      <div className="lookupCriteria">
        <div className="lookupHeader">
          <strong>Match Criteria</strong>
          <button type="button" onClick={addInput}>Add Criteria</button>
        </div>
        {inputs.map((input, index) => (
          <div className="lookupCriterion" key={`${input.column}-${index}`}>
            <label>
              Table Column
              <input value={input.column} onChange={(event) => updateInput(index, { column: sanitizeName(event.target.value) })} />
            </label>
            <label>
              Smart Cell Reference
              <input value={input.reference} onChange={(event) => updateInput(index, { reference: sanitizeName(event.target.value) })} />
            </label>
            <button type="button" onClick={() => removeInput(index)} disabled={inputs.length <= 1}>Remove</button>
          </div>
        ))}
      </div>

      <label>
        Output Column
        <input value={lookup.outputColumn} onChange={(event) => updateLookup({ outputColumn: sanitizeName(event.target.value) })} />
      </label>

      <div className="lookupTableScroll">
        <div className="lookupRows">
          <div className="lookupRow lookupRowHeader" style={{ gridTemplateColumns: `repeat(${tableColumns.length}, minmax(110px, 1fr))` }}>
            {tableColumns.map((column) => (
              <span key={column}>{column}</span>
            ))}
          </div>
          {lookup.rows.map((row, index) => (
            <div className="lookupRow" key={index} style={{ gridTemplateColumns: `repeat(${tableColumns.length}, minmax(110px, 1fr))` }}>
              {tableColumns.map((column) => (
                <input
                  key={column}
                  value={String(row[column] ?? "")}
                  onChange={(event) => updateRow(index, column, event.target.value)}
                  onPaste={(event) => handlePaste(event, index, column)}
                />
              ))}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function makeCell(
  address: string,
  entry: string,
  type: SmartCellType,
  options: Partial<Omit<GridCell, "address" | "entry" | "type">> = {},
): GridCell {
  const cell: GridCell = {
    address,
    entry,
    name: "",
    label: "",
    role: entry.startsWith("=") ? "formula" : "input",
    type,
    inputControl: "freeText",
    inputOptions: [],
    surfaced: false,
    runnerSection: "",
    annotation: "",
    ruleMessage: "",
    ...options,
  };

  if (cell.inputOptions.length > 0 && options.inputControl === undefined) {
    cell.inputControl = "dropdown";
  }

  return cell;
}

function hydrateCells(cells: Record<string, GridCell>): Record<string, GridCell> {
  return Object.fromEntries(
    Object.entries(cells).map(([address, cell]) => {
      const inputOptions = cell.inputOptions ?? [];
      const hydrated = {
        ...makeCell(address, "", "text"),
        ...cell,
        inputControl: cell.inputControl ?? (inputOptions.length > 0 ? "dropdown" : "freeText"),
        inputOptions,
        runnerSection: cell.runnerSection ?? "",
        lookup: cell.lookup ?? (cell.role === "lookup" || cell.role === "action" ? starterLookup : undefined),
      };

      return [address, applyCellPatch(hydrated, {})];
    }),
  );
}

function runnerOverrideKey(sheetId: string, address: string): string {
  return `${sheetId}!${address}`;
}

function applyRunnerOverridesToSheets(sheets: WorkbookSheet[], overrides: RunnerOverrides): WorkbookSheet[] {
  return sheets.map((sheet) => {
    let cells = sheet.cells;

    for (const [key, entry] of Object.entries(overrides)) {
      if (!key.startsWith(`${sheet.id}!`)) continue;
      const address = key.slice(sheet.id.length + 1);
      const existing = getCell(cells, address);
      cells = {
        ...cells,
        [address]: applyCellPatch(existing, { entry }),
      };
    }

    return cells === sheet.cells ? sheet : { ...sheet, cells };
  });
}

function cellsEqual(left: Record<string, GridCell>, right: Record<string, GridCell>): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function buildDependencySummary(cells: Record<string, GridCell>, selectedAddress: string): DependencySummary {
  const selectedCell = getCell(cells, selectedAddress);
  const referenceIndex = buildReferenceIndex(cells);
  const selectedReferences = referencesForAddress(selectedCell);
  const dependencyItems = selectedReferences
    .map((reference) => {
      const dependency = referenceIndex.get(reference);
      if (!dependency) return null;
      return dependencyItem(dependency, reference);
    })
    .filter(Boolean) as DependencyItem[];

  const selectedKeys = new Set([selectedAddress, selectedCell.name].filter(Boolean));
  const dependentItems: DependencyItem[] = [];

  for (const cell of Object.values(cells)) {
    if (cell.address === selectedAddress) continue;
    const refs = referencesForAddress(cell);
    if (!refs.some((reference) => selectedKeys.has(reference))) continue;
    dependentItems.push(dependencyItem(cell, cell.name || cell.address));
  }

  return {
    dependencies: dedupeDependencyItems(dependencyItems),
    dependents: dedupeDependencyItems(dependentItems),
  };
}

function buildReferenceIndex(cells: Record<string, GridCell>): Map<string, GridCell> {
  const index = new Map<string, GridCell>();

  for (const cell of Object.values(cells)) {
    index.set(cell.address, cell);
    if (cell.name) index.set(cell.name, cell);
  }

  return index;
}

function referencesForAddress(cell: GridCell): string[] {
  const refs = new Set<string>();
  if (cell.entry.trim().startsWith("=")) {
    for (const reference of referencesForFormula(cell.entry)) refs.add(reference);
  }

  if ((cell.role === "lookup" || cell.role === "action") && cell.lookup) {
    for (const input of normalizeLookupInputs(cell.lookup)) {
      if (input.reference) refs.add(input.reference);
    }
  }

  if ((cell.role === "validation" || cell.role === "compliance") && cell.entry) {
    for (const reference of referencesForFormula(cell.entry)) refs.add(reference);
  }

  return [...refs];
}

function referencesForFormula(entry: string): string[] {
  const refs = new Set<string>();
  for (const reference of parseFormula(entry).references) {
    if (reference.sheetName) continue;
    if (reference.kind === "name") refs.add(reference.raw);
    if (reference.kind === "cell" && reference.address) refs.add(reference.address.replace(/\$/g, "").toUpperCase());
    if (reference.kind === "range" && reference.address && reference.rangeEnd) {
      for (const address of expandAddressRange(reference.address.replace(/\$/g, ""), reference.rangeEnd.replace(/\$/g, ""))) refs.add(address);
    }
  }
  return [...refs];
}

function expandAddressRange(start: string, end: string): string[] {
  const startRef = parseAddress(start.toUpperCase());
  const endRef = parseAddress(end.toUpperCase());
  if (!startRef || !endRef) return [];

  const firstColumn = Math.min(columnNumber(startRef.column), columnNumber(endRef.column));
  const lastColumn = Math.max(columnNumber(startRef.column), columnNumber(endRef.column));
  const firstRow = Math.min(startRef.row, endRef.row);
  const lastRow = Math.max(startRef.row, endRef.row);
  const addresses: string[] = [];

  for (let row = firstRow; row <= lastRow; row++) {
    for (let column = firstColumn; column <= lastColumn; column++) {
      addresses.push(`${columnName(column) ?? "A"}${row}`);
    }
  }

  return addresses;
}

function dependencyItem(cell: GridCell, reference: string): DependencyItem {
  return {
    address: cell.address,
    label: labelForCell(cell),
    reference,
  };
}

function dedupeDependencyItems(items: DependencyItem[]): DependencyItem[] {
  const seen = new Set<string>();
  return items.filter((item) => {
    const key = `${item.address}:${item.reference}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function insertRow(cells: Record<string, GridCell>, insertAt: number): Record<string, GridCell> {
  return transformSheet(cells, ({ column, row, cell }) => {
    const nextRow = row >= insertAt ? row + 1 : row;
    return {
      address: `${column}${nextRow}`,
      cell: updateCellAddress(cell, `${column}${nextRow}`, (entry) => shiftInsertedRowReferences(entry, insertAt)),
    };
  });
}

function deleteRow(cells: Record<string, GridCell>, deleteAt: number): Record<string, GridCell> {
  return transformSheet(cells, ({ column, row, cell }) => {
    if (row === deleteAt) return null;
    const nextRow = row > deleteAt ? row - 1 : row;
    return {
      address: `${column}${nextRow}`,
      cell: updateCellAddress(cell, `${column}${nextRow}`, (entry) => shiftDeletedRowReferences(entry, deleteAt)),
    };
  });
}

function insertColumn(cells: Record<string, GridCell>, insertAt: number): Record<string, GridCell> {
  return transformSheet(cells, ({ column, row, cell }) => {
    const columnIndex = columnNumber(column);
    const nextColumn = columnName(columnIndex >= insertAt ? columnIndex + 1 : columnIndex) ?? column;
    return {
      address: `${nextColumn}${row}`,
      cell: updateCellAddress(cell, `${nextColumn}${row}`, (entry) => shiftInsertedColumnReferences(entry, insertAt)),
    };
  });
}

function deleteColumn(cells: Record<string, GridCell>, deleteAt: number): Record<string, GridCell> {
  return transformSheet(cells, ({ column, row, cell }) => {
    const columnIndex = columnNumber(column);
    if (columnIndex === deleteAt) return null;
    const nextColumn = columnName(columnIndex > deleteAt ? columnIndex - 1 : columnIndex) ?? column;
    return {
      address: `${nextColumn}${row}`,
      cell: updateCellAddress(cell, `${nextColumn}${row}`, (entry) => shiftDeletedColumnReferences(entry, deleteAt)),
    };
  });
}

function transformSheet(
  cells: Record<string, GridCell>,
  transform: (input: { column: string; row: number; cell: GridCell }) => { address: string; cell: GridCell } | null,
): Record<string, GridCell> {
  const next: Record<string, GridCell> = {};

  for (const [address, cell] of Object.entries(cells)) {
    const parsed = parseAddress(address);
    if (!parsed) {
      next[address] = cell;
      continue;
    }

    const transformed = transform({ ...parsed, cell });
    if (!transformed) continue;
    next[transformed.address] = transformed.cell;
  }

  return next;
}

function updateCellAddress(cell: GridCell, address: string, formulaTransform: (entry: string) => string): GridCell {
  return {
    ...cell,
    address,
    entry: cell.entry.trim().startsWith("=") ? formulaTransform(cell.entry) : cell.entry,
  };
}

function renameSheetReferences(cells: Record<string, GridCell>, oldName: string, newName: string): Record<string, GridCell> {
  return Object.fromEntries(
    Object.entries(cells).map(([address, cell]) => [
      address,
      {
        ...cell,
        entry: replaceSheetNameInFormula(cell.entry, oldName, newName),
        lookup: cell.lookup
          ? {
              ...cell.lookup,
              inputReference: replaceSheetNameInFormula(cell.lookup.inputReference, oldName, newName),
              inputs: cell.lookup.inputs?.map((input) => ({
                ...input,
                reference: replaceSheetNameInFormula(input.reference, oldName, newName),
              })),
            }
          : cell.lookup,
      },
    ]),
  );
}

function replaceSheetNameInFormula(entry: string, oldName: string, newName: string): string {
  if (!entry.includes("!")) return entry;
  return rewriteFormula(entry, (reference) => {
    if (reference.sheetName?.toLowerCase() !== oldName.toLowerCase()) return null;
    const start = reference.address ?? "";
    const range = reference.rangeEnd ? `:${reference.rangeEnd}` : "";
    return `${quoteFormulaSheetName(newName)}!${start}${range}`;
  });
}

function shiftInsertedRowReferences(entry: string, insertAt: number): string {
  return replaceFormulaReferences(entry, (column, row) => `${column}${row >= insertAt ? row + 1 : row}`);
}

function shiftDeletedRowReferences(entry: string, deleteAt: number): string {
  return replaceFormulaReferences(entry, (column, row) => {
    if (row === deleteAt) return "#REF!";
    return `${column}${row > deleteAt ? row - 1 : row}`;
  });
}

function shiftInsertedColumnReferences(entry: string, insertAt: number): string {
  return replaceFormulaReferences(entry, (column, row) => {
    const columnIndex = columnNumber(column);
    return `${columnName(columnIndex >= insertAt ? columnIndex + 1 : columnIndex) ?? column}${row}`;
  });
}

function shiftDeletedColumnReferences(entry: string, deleteAt: number): string {
  return replaceFormulaReferences(entry, (column, row) => {
    const columnIndex = columnNumber(column);
    if (columnIndex === deleteAt) return "#REF!";
    return `${columnName(columnIndex > deleteAt ? columnIndex - 1 : columnIndex) ?? column}${row}`;
  });
}

function replaceFormulaReferences(entry: string, replacer: (column: string, row: number) => string): string {
  return rewriteFormula(entry, (reference) => {
    if (reference.sheetName || reference.kind === "name" || !reference.address) return null;
    const replaceAddress = (address: string) => {
      const parsed = parseAddress(address.replace(/\$/g, ""));
      return parsed ? replacer(parsed.column, parsed.row) : address;
    };
    const start = replaceAddress(reference.address);
    return reference.rangeEnd ? `${start}:${replaceAddress(reference.rangeEnd)}` : start;
  });
}

function normalizeLookupInputs(lookup: LookupConfig): Array<{ column: string; reference: string }> {
  if (lookup.inputs?.length) return lookup.inputs;
  return [{ column: lookup.inputColumn, reference: lookup.inputReference }];
}

function hydrateConfigurations(configurations: LocalConfiguration[]): LocalConfiguration[] {
  if (!Array.isArray(configurations)) return [];

  return configurations
    .filter((configuration) => configuration && typeof configuration === "object")
    .map((configuration) => {
      const legacySheet = sheetFromConfiguration(configuration);
      const sheets = Array.isArray(configuration.sheets) && configuration.sheets.length > 0
        ? configuration.sheets.map(hydrateWorkbookSheet)
        : [legacySheet];
      const activeSheet = sheets.find((sheet) => sheet.id === configuration.activeSheetId) ?? sheets[0];

      return {
        id: configuration.id || makeConfigId(),
        name: configuration.name || "Untitled Configuration",
        activeSheetId: activeSheet.id,
        sheets,
        updatedAt: configuration.updatedAt || new Date().toISOString(),
      };
    });
}

function migrateLegacyCells(): Record<string, GridCell> | null {
  try {
    const stored = window.localStorage.getItem(STORAGE_KEY);
    if (!stored) return null;

    const parsed = JSON.parse(stored) as Record<string, GridCell>;
    if (!parsed || typeof parsed !== "object") return null;

    return hydrateCells(parsed);
  } catch {
    return null;
  }
}

function makeConfiguration(
  name: string,
  cells: Record<string, GridCell>,
  columnCount = defaultColumnCount,
  rowCount = defaultRowCount,
  workbook?: { sheets: WorkbookSheet[]; activeSheetId: string },
): LocalConfiguration {
  const sheets = workbook?.sheets.map(hydrateWorkbookSheet)
    ?? [makeWorkbookSheet("Sheet 1", cells, columnCount, rowCount)];
  const activeSheet = sheets.find((sheet) => sheet.id === workbook?.activeSheetId) ?? sheets[0];
  return {
    id: makeConfigId(),
    name,
    activeSheetId: activeSheet?.id,
    sheets,
    updatedAt: new Date().toISOString(),
  };
}

function makeImportedConfigurationName(fileName: string): string {
  const baseName = fileName.replace(/\.[^.]+$/, "").trim() || "Workbook";
  return `Imported - ${baseName}`;
}

function importReviewItemsForSheet(names: ImportedName[], sheetName: string): ImportReviewItem[] {
  return names
    .filter((name) => !name.sheetName || name.sheetName === sheetName)
    .filter((name) => name.kind !== "singleCell")
    .map((name) => ({
      severity: "info",
      sheetName: name.sheetName ?? sheetName,
      name: name.name,
      message: `Workbook name "${name.name}" points to ${name.kind}; it will be reported for review.`,
    }));
}

function makeConfigId(): string {
  return `config_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function makeSheetId(): string {
  return `sheet_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function makeWorkbookSheet(
  name: string,
  cells: Record<string, GridCell>,
  columnCount = defaultColumnCount,
  rowCount = defaultRowCount,
): WorkbookSheet {
  return {
    id: makeSheetId(),
    name: name.trim() || "Sheet",
    cells: hydrateCells(cells),
    columnCount: Math.max(columnCount, defaultColumnCount),
    rowCount: Math.max(rowCount, defaultRowCount),
  };
}

function hydrateWorkbookSheet(sheet: WorkbookSheet): WorkbookSheet {
  return {
    id: sheet.id || makeSheetId(),
    name: sheet.name || "Sheet",
    cells: hydrateCells(sheet.cells ?? {}),
    columnCount: Math.max(sheet.columnCount ?? defaultColumnCount, defaultColumnCount),
    rowCount: Math.max(sheet.rowCount ?? defaultRowCount, defaultRowCount),
  };
}

function sheetFromConfiguration(configuration: LocalConfiguration): WorkbookSheet {
  return {
    id: configuration.activeSheetId || makeSheetId(),
    name: "Sheet 1",
    cells: hydrateCells(configuration.cells ?? {}),
    columnCount: Math.max(configuration.columnCount ?? defaultColumnCount, defaultColumnCount),
    rowCount: Math.max(configuration.rowCount ?? defaultRowCount, defaultRowCount),
  };
}

function getCell(cells: Record<string, GridCell>, address: string): GridCell {
  return cells[address] ?? makeCell(address, "", "text");
}

function applyCellPatch(cell: GridCell, patch: Partial<GridCell>): GridCell {
  const next = { ...cell, ...patch };
  const hasName = Boolean(next.name);
  const hasFormulaEntry = typeof next.entry === "string" && next.entry.trim().startsWith("=");
  next.inputOptions = next.inputOptions ?? [];
  next.inputControl = next.inputControl ?? (next.inputOptions.length > 0 ? "dropdown" : "freeText");

  if (patch.entry !== undefined && patch.type === undefined && !cell.name) {
    next.type = inferPrimitiveEntryType(next.entry, next.type);
  }

  if (hasName && hasFormulaEntry && next.role === "input" && patch.role === undefined) {
    next.role = "formula";
  }

  return next;
}

function isDropdownCell(cell: GridCell): boolean {
  return cell.role === "input" && cell.inputControl === "dropdown";
}

function isCheckboxCell(cell: GridCell): boolean {
  return cell.role === "input" && cell.inputControl === "checkbox";
}

function dropdownOptionsForCell(cell: GridCell): string[] {
  if (!isDropdownCell(cell)) return [];
  if (cell.entry && !cell.inputOptions.includes(cell.entry)) return [cell.entry, ...cell.inputOptions];
  return cell.inputOptions;
}

function inferPrimitiveEntryType(entry: string, fallback: SmartCellType): SmartCellType {
  const trimmed = entry.trim();
  if (!trimmed || trimmed.startsWith("=")) return fallback;
  if (/^(true|false)$/i.test(trimmed)) return "boolean";
  const numericValue = Number(trimmed);
  return Number.isFinite(numericValue) ? "number" : "text";
}

function toEngineCells(cells: Record<string, GridCell>): EngineCell[] {
  return Object.values(cells)
    .filter((cell) => cell.entry !== "" || cell.name)
    .map((cell) => {
      const isFormula = cell.entry.trim().startsWith("=");
      const engineCell: EngineCell = {
        id: cell.address,
        address: cell.address,
        name: cell.name || null,
        role: cell.name ? cell.role : isFormula ? "formula" : "input",
        type: cell.type,
        value: isFormula ? null : parseCellValue(cell.entry, cell.type),
        formula: isFormula ? normalizeFormula(cell.entry) : null,
        annotation: cell.annotation || null,
        surfaced: Boolean(cell.name && cell.surfaced),
      };

      if (cell.name && (cell.role === "lookup" || cell.role === "action") && cell.lookup) {
        const lookup = cell.lookup ?? starterLookup;
        engineCell.formula = null;
        engineCell.lookup = {
          inputMap: Object.fromEntries(normalizeLookupInputs(lookup).map((input) => [input.column, input.reference])),
          outputColumn: lookup.outputColumn,
          rows: lookup.rows,
        };
      }

      if (cell.name && cell.role === "validation") {
        engineCell.validation = {
          condition: normalizeFormula(cell.entry) || "true",
          message: cell.ruleMessage || cell.annotation || "Validation failed.",
        };
      }

      if (cell.name && cell.role === "compliance") {
        engineCell.compliance = {
          condition: normalizeFormula(cell.entry) || "false",
          message: cell.ruleMessage || cell.annotation || "Compliance warning.",
        };
      }

      return engineCell;
    });
}

function buildRunnerSheetContexts(sheets: WorkbookSheet[], workbookResult: WorkbookEngineResult): RunnerSheetContext[] {
  const visibilityValues = buildWorkbookVisibilityValues(sheets, workbookResult);
  return sheets.map((sheet) => {
    const sheetResult = workbookResult.sheetResults.find((item) => item.sheetId === sheet.id);
    const resultForSheet = sheetResult?.result ?? executeEngine({ cells: toEngineCells(sheet.cells) });
    const ruleStateMap = new Map(resultForSheet.ruleStates.map((rule) => [rule.address, rule.state]));
    const columns = makeColumns(sheet.columnCount);
    const displayValues = buildDisplayValues(sheet.cells, resultForSheet.values, resultForSheet.errors, ruleStateMap, columns, sheet.rowCount);
    const surfacedCells = Object.values(sheet.cells)
      .filter((cell) => cell.name && cell.surfaced)
      .filter((cell) => isCellVisibleInRunner(cell, visibilityValues));
    const validationStates = resultForSheet.ruleStates.filter((rule) => {
      const cell = getCell(sheet.cells, rule.address);
      return cell.role === "validation" && cell.surfaced && isCellVisibleInRunner(cell, visibilityValues);
    });

    return {
      sheetId: sheet.id,
      sheetName: sheet.name,
      cells: sheet.cells,
      displayValues,
      surfacedCells,
      result: resultForSheet,
      validationStates,
    };
  });
}

function buildWorkbookVisibilityValues(sheets: WorkbookSheet[], workbookResult: WorkbookEngineResult): Map<string, CellValue> {
  const values = new Map<string, CellValue>();

  for (const sheet of sheets) {
    const sheetResult = workbookResult.sheetResults.find((item) => item.sheetId === sheet.id);
    const resultForSheet = sheetResult?.result ?? executeEngine({ cells: toEngineCells(sheet.cells) });

    for (const cell of Object.values(sheet.cells)) {
      if (!cell.name) continue;
      const value = resultForSheet.values[cell.address] ?? parseCellValue(cell.entry, cell.type);
      if (!values.has(cell.name)) values.set(cell.name, value);
    }
  }

  return values;
}

function isCellVisibleInRunner(cell: GridCell, visibilityValues: Map<string, CellValue>): boolean {
  if (!cell.visibilityCondition) return true;
  return isVisibilityConditionMet(cell.visibilityCondition, visibilityValues);
}

function isVisibilityConditionMet(condition: SmartCellVisibilityCondition, visibilityValues: Map<string, CellValue>): boolean {
  if (!condition.source || !visibilityValues.has(condition.source)) return false;
  const sourceValue = visibilityValues.get(condition.source);
  if (sourceValue === undefined) return false;
  return normalizeConditionValue(sourceValue) === normalizeConditionValue(condition.value);
}

function normalizeConditionValue(value: CellValue): string {
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "number") return String(value);
  return String(value ?? "").trim().toLowerCase();
}

function buildVisibilityControlOptions(sheets: WorkbookSheet[]): VisibilityControlOption[] {
  const names = new Map<string, number>();
  for (const sheet of sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (cell.name) names.set(cell.name, (names.get(cell.name) ?? 0) + 1);
    }
  }

  return sheets.flatMap((sheet) => Object.values(sheet.cells)
    .filter((cell) => cell.name && names.get(cell.name) === 1 && cell.role === "input" && cell.type === "boolean")
    .map((cell) => ({
      name: cell.name,
      label: labelForCell(cell),
      sheetName: sheet.name,
    })));
}

function renameVisibilityConditionReferences(cells: Record<string, GridCell>, oldName: string, newName: string): Record<string, GridCell> {
  let changed = false;
  const next = Object.fromEntries(Object.entries(cells).map(([address, cell]) => {
    if (cell.visibilityCondition?.source !== oldName) return [address, cell];
    changed = true;
    return [address, {
      ...cell,
      visibilityCondition: {
        ...cell.visibilityCondition,
        source: newName,
      },
    }];
  })) as Record<string, GridCell>;

  return changed ? next : cells;
}

function groupItemsByRunnerSection<T>(items: T[], cellForItem: (item: T) => GridCell): Array<{ label: string; items: T[] }> {
  const groups: Array<{ label: string; items: T[] }> = [];

  for (const item of items) {
    const label = cellForItem(item).runnerSection.trim();
    let group = groups.find((candidate) => candidate.label === label);
    if (!group) {
      group = { label, items: [] };
      groups.push(group);
    }
    group.items.push(item);
  }

  return groups;
}

function buildDisplayValues(
  cells: Record<string, GridCell>,
  values: Record<string, CellValue>,
  errors: Array<{ address: string }>,
  ruleStateMap: Map<string, string>,
  columns: string[],
  rowCount: number,
): Record<string, CellValue> {
  const display: Record<string, CellValue> = {};
  const errorAddresses = new Set(errors.map((error) => error.address));

  for (let row = 1; row <= rowCount; row++) {
    for (const column of columns) {
      const address = `${column}${row}`;
      const cell = getCell(cells, address);
      const isFormula = cell.entry.trim().startsWith("=");
      const ruleState = ruleStateMap.get(address);
      if (cell.role === "validation") {
        display[address] = ruleState === "fail" ? "FAIL" : ruleState === "error" ? "#ERR" : ruleState === "ok" ? "PASS" : "";
      } else if (cell.role === "compliance") {
        display[address] = ruleState === "warn" ? "WARN" : ruleState === "error" ? "#ERR" : ruleState === "ok" ? "OK" : "";
      } else if (isFormula || cell.role === "lookup" || cell.role === "action") {
        display[address] = errorAddresses.has(address) ? "#ERR" : values[cell.name || address] ?? "";
      } else {
        display[address] = parseCellValue(cell.entry, cell.type);
      }
    }
  }

  return display;
}

function buildColumnWidths(
  cells: Record<string, GridCell>,
  displayValues: Record<string, CellValue>,
  columns: string[],
  rowCount: number,
): number[] {
  return columns.map((column) => {
    let maxLength = column.length;

    for (let row = 1; row <= rowCount; row += 1) {
      const address = `${column}${row}`;
      const cell = getCell(cells, address);
      const visibleValue = formatCellValue(displayValues[address] ?? null);
      const marker = cell.name ? cell.name.length + 2 : 0;
      maxLength = Math.max(maxLength, visibleValue.length, marker);
    }

    return Math.min(220, Math.max(64, maxLength * 6 + 18));
  });
}

function buildWorkbookReferenceOptions(sheets: WorkbookSheet[], workbookResult: WorkbookEngineResult, activeSheetId: string): ReferenceOption[] {
  const options: ReferenceOption[] = [];
  const nameCounts = new Map<string, number>();

  for (const sheet of sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (!cell.name) continue;
      nameCounts.set(cell.name, (nameCounts.get(cell.name) ?? 0) + 1);
    }
  }

  for (const sheet of sheets) {
    const resultForSheet = workbookResult.sheetResults.find((item) => item.sheetId === sheet.id)?.result ?? executeEngine({ cells: toEngineCells(sheet.cells) });
    const ruleStateMap = new Map(resultForSheet.ruleStates.map((rule) => [rule.address, rule.state]));
    const columns = makeColumns(sheet.columnCount);
    const displayValues = buildDisplayValues(sheet.cells, resultForSheet.values, resultForSheet.errors, ruleStateMap, columns, sheet.rowCount);

    for (const option of buildSheetReferenceOptions(sheet, displayValues, nameCounts, sheet.id === activeSheetId)) {
      options.push(option);
    }
  }

  return options;
}

function buildSheetReferenceOptions(sheet: WorkbookSheet, displayValues: Record<string, CellValue>, nameCounts: Map<string, number>, isActiveSheet: boolean): ReferenceOption[] {
  const options: ReferenceOption[] = [];
  const columns = makeColumns(sheet.columnCount);

  for (let row = 1; row <= sheet.rowCount; row++) {
    for (const column of columns) {
      const address = `${column}${row}`;
      const cell = getCell(sheet.cells, address);
      if (!cell.entry && !cell.name) continue;
      const canUseWorkbookName = Boolean(cell.name && nameCounts.get(cell.name) === 1);
      options.push({
        address: `${sheet.name}!${address}`,
        key: referenceOptionKey(sheet.id, address),
        reference: canUseWorkbookName ? cell.name : isActiveSheet ? address : `${quoteSheetName(sheet.name)}!${address}`,
        value: displayValues[address] ?? "",
      });
    }
  }

  return options;
}

function referenceOptionKey(sheetId: string, address: string): string {
  return `${sheetId}!${address}`;
}

function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function normalizeFormula(entry: string): string {
  const trimmed = entry.trim();
  if (!trimmed.startsWith("=")) return trimmed;
  return trimmed.slice(1);
}

function parseCellValue(value: string, type: SmartCellType): CellValue {
  if (value.trim() === "") return null;
  if (type === "number") {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  if (type === "boolean") return value === "true" || value === "1" || value.toLowerCase() === "yes";
  return value;
}

function parseLooseValue(value: string): CellValue {
  const numeric = Number(value);
  if (value.trim() !== "" && Number.isFinite(numeric)) return numeric;
  if (value === "true") return true;
  if (value === "false") return false;
  return value;
}

function parsePastedRows(text: string): string[][] {
  return text
    .replace(/\r/g, "")
    .split("\n")
    .filter((row, index, rows) => row !== "" || index < rows.length - 1)
    .map((row) => row.split("\t"));
}

function firstRowLooksLikeLookupHeader(row: string[], tableColumns: string[]): boolean {
  if (row.length === 0) return false;
  return row.every((cell, index) => {
    const expected = tableColumns[index];
    return expected && normalizeHeader(cell) === normalizeHeader(expected);
  });
}

function normalizeHeader(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, "_");
}

function splitInputOptions(value: string): string[] {
  return value
    .split(/\r?\n|,/)
    .map((option) => option.trim())
    .filter(Boolean);
}

function adjustFormulaReferences(entry: string, rowOffset: number, columnOffset: number): string {
  return rewriteFormula(entry, (reference) => {
    if (reference.sheetName || reference.kind === "name" || !reference.address) return null;
    const adjust = (address: string) => {
      const parsed = parseAddress(address.replace(/\$/g, ""));
      if (!parsed) return address;
      const nextColumn = columnName(columnNumber(parsed.column) + columnOffset);
      const nextRow = parsed.row + rowOffset;
      return !nextColumn || nextRow < 1 ? address : `${nextColumn}${nextRow}`;
    };
    const start = adjust(reference.address);
    return reference.rangeEnd ? `${start}:${adjust(reference.rangeEnd)}` : start;
  });
}

function columnNumber(column: string): number {
  return column.split("").reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function columnName(column: number): string | null {
  if (column < 1) return null;

  let remaining = column;
  let name = "";

  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    remaining = Math.floor((remaining - modulo) / 26);
  }

  return name;
}

function makeColumns(count: number): string[] {
  return Array.from({ length: count }, (_, index) => columnName(index + 1) ?? "A");
}

function parseAddress(address: string): { column: string; row: number } | null {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  if (!match) return null;
  return { column: match[1], row: Number(match[2]) };
}

function getReferenceQuery(entry: string): string {
  const tokenStart = getReferenceTokenStart(entry);
  if (tokenStart === null) return "";
  return entry.slice(tokenStart);
}

function getReferenceTokenStart(entry: string): number | null {
  if (!entry.trim().startsWith("=")) return null;

  for (let index = entry.length - 1; index >= 0; index--) {
    if (/[\s=+\-*/(),<>]/.test(entry[index])) return index + 1;
  }

  return 1;
}

function sanitizeName(value: string): string {
  return value.replace(/[^a-zA-Z0-9_]/g, "_");
}

function labelForCell(cell: GridCell): string {
  return cell.label.trim() || (cell.name ? prettifyName(cell.name) : cell.address);
}

function labelForInputControl(inputControl: InputControl): string {
  if (inputControl === "dropdown") return "Dropdown";
  if (inputControl === "checkbox") return "Checkbox";
  return "Free text";
}

function formatCellValue(value: CellValue): string {
  if (value === null) return "";
  return String(value);
}

function prettifyName(name: string): string {
  return name
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}
