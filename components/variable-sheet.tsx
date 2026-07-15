"use client";

import { useEffect, useMemo, useRef, useState, type ChangeEvent, type ReactNode } from "react";
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
import type { ImportedName, ImportedWorkbook, ImportReviewItem } from "@/lib/import/types";
import type { GridCell, InputControl, LocalConfiguration, LookupConfig, SheetSnapshot, WorkbookSheet } from "@/lib/sheet/types";

const STORAGE_KEY = "quoin.gridSheet.v2";
const CONFIG_STORAGE_KEY = "quoin.configurations.v1";
const ACTIVE_CONFIG_KEY = "quoin.activeConfiguration.v1";
const defaultColumnCount = 16;
const defaultRowCount = 30;
const historyLimit = 50;
const roleOptions: SmartCellRole[] = ["input", "formula", "output", "action", "lookup", "validation", "compliance"];
const typeOptions: SmartCellType[] = ["number", "text", "boolean"];
const inputControlOptions: InputControl[] = ["freeText", "dropdown"];

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
  const [sheets, setSheets] = useState<WorkbookSheet[]>(() => [makeWorkbookSheet("Sheet 1", initialCells, defaultColumnCount, defaultRowCount)]);
  const [activeSheetId, setActiveSheetId] = useState("");
  const [cells, setCells] = useState<Record<string, GridCell>>(initialCells);
  const [columnCount, setColumnCount] = useState(defaultColumnCount);
  const [rowCount, setRowCount] = useState(defaultRowCount);
  const [configurations, setConfigurations] = useState<LocalConfiguration[]>([]);
  const [activeConfigId, setActiveConfigId] = useState("");
  const [configName, setConfigName] = useState("Demo - Beam Selection");
  const [isDirty, setIsDirty] = useState(false);
  const [selectedAddress, setSelectedAddress] = useState("B2");
  const [editingAddress, setEditingAddress] = useState<string | null>(null);
  const [draftEntry, setDraftEntry] = useState("");
  const [isLoaded, setIsLoaded] = useState(false);
  const [activeView, setActiveView] = useState<"sheet" | "runner" | "help">("sheet");
  const [activeReferenceIndex, setActiveReferenceIndex] = useState(0);
  const [undoStack, setUndoStack] = useState<SheetSnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<SheetSnapshot[]>([]);
  const [copiedAddress, setCopiedAddress] = useState<string | null>(null);
  const [pendingImport, setPendingImport] = useState<ImportedWorkbook | null>(null);
  const [selectedImportSheetId, setSelectedImportSheetId] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [importMessage, setImportMessage] = useState("");
  const [importError, setImportError] = useState("");
  const [dropdownOptionsDraft, setDropdownOptionsDraft] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);
  const cellRefs = useRef<Record<string, HTMLDivElement | null>>({});

  const activeSheet = useMemo(() => sheets.find((sheet) => sheet.id === activeSheetId) ?? sheets[0] ?? null, [activeSheetId, sheets]);
  const columns = useMemo(() => makeColumns(columnCount), [columnCount]);
  const selectedCell = getCell(cells, selectedAddress);
  const selectedDropdownOptionsText = selectedCell.inputOptions.join("\n");
  const engineCells = useMemo(() => toEngineCells(cells), [cells]);
  const result = useMemo(() => executeEngine({ cells: engineCells }), [engineCells]);
  const ruleStateMap = useMemo(() => new Map(result.ruleStates.map((rule) => [rule.address, rule.state])), [result.ruleStates]);
  const displayValues = useMemo(
    () => buildDisplayValues(cells, result.values, result.errors, ruleStateMap, columns, rowCount),
    [cells, columns, result.errors, result.values, rowCount, ruleStateMap],
  );
  const columnWidths = useMemo(() => buildColumnWidths(cells, displayValues, columns, rowCount), [cells, columns, displayValues, rowCount]);
  const referenceOptions = useMemo(() => buildReferenceOptions(cells, displayValues, columns, rowCount), [cells, columns, displayValues, rowCount]);
  const filteredReferenceOptions = useMemo(() => {
    const query = getReferenceQuery(draftEntry);
    const normalized = query.toLowerCase();
    if (!normalized) return referenceOptions;

    return referenceOptions.filter((option) => {
      return option.reference.toLowerCase().includes(normalized) || option.address.toLowerCase().includes(normalized);
    });
  }, [draftEntry, referenceOptions]);
  const visibleReferenceOptions = useMemo(() => {
    return filteredReferenceOptions.filter((option) => option.address !== editingAddress).slice(0, 10);
  }, [editingAddress, filteredReferenceOptions]);
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
    setDropdownOptionsDraft(selectedDropdownOptionsText);
  }, [selectedAddress, selectedDropdownOptionsText]);

  useEffect(() => {
    try {
      const storedConfigs = window.localStorage.getItem(CONFIG_STORAGE_KEY);
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
      setSheets(activeConfig.sheets ?? [sheetFromConfiguration(activeConfig)]);
      setActiveSheetId(activeConfig.activeSheetId ?? activeConfig.sheets?.[0]?.id ?? "");
      setCells(activeConfig.cells);
      setColumnCount(Math.max(activeConfig.columnCount ?? defaultColumnCount, defaultColumnCount));
      setRowCount(Math.max(activeConfig.rowCount ?? defaultRowCount, defaultRowCount));
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
    editInputRef.current?.focus();
    editInputRef.current?.select();
  }, [editingAddress]);

  useEffect(() => {
    if (editingAddress || activeView !== "sheet") return;
    cellRefs.current[selectedAddress]?.focus();
  }, [activeView, editingAddress, selectedAddress]);

  useEffect(() => {
    setActiveReferenceIndex(0);
  }, [draftEntry, editingAddress]);

  useEffect(() => {
    setActiveReferenceIndex((current) => Math.min(current, Math.max(0, visibleReferenceOptions.length - 1)));
  }, [visibleReferenceOptions.length]);

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

  function applyCellsChange(updater: (current: Record<string, GridCell>) => Record<string, GridCell>) {
    setCells((current) => {
      const next = updater(current);
      if (cellsEqual(current, next)) return current;

      setUndoStack((history) => [...history.slice(Math.max(0, history.length - historyLimit + 1)), makeSnapshot(current, columnCount, rowCount)]);
      setRedoStack([]);
      return next;
    });
    setIsDirty(true);
  }

  function applySheetResize(updater: (current: Record<string, GridCell>) => Record<string, GridCell>, nextColumnCount: number, nextRowCount: number) {
    setUndoStack((history) => [...history.slice(Math.max(0, history.length - historyLimit + 1)), makeSnapshot(cells, columnCount, rowCount)]);
    setRedoStack([]);
    setCells((current) => updater(current));
    setColumnCount(nextColumnCount);
    setRowCount(nextRowCount);
    setIsDirty(true);
  }

  function updateCell(address: string, patch: Partial<GridCell>) {
    applyCellsChange((current) => {
      const existing = getCell(current, address);
      return { ...current, [address]: applyCellPatch(existing, patch) };
    });
  }

  function updateWorkbookCell(sheetId: string, address: string, patch: Partial<GridCell>) {
    if (sheetId === activeSheetId) {
      updateCell(address, patch);
      return;
    }

    setSheets((current) => current.map((sheet) => {
      if (sheet.id !== sheetId) return sheet;
      const existing = getCell(sheet.cells, address);
      return {
        ...sheet,
        cells: {
          ...sheet.cells,
          [address]: applyCellPatch(existing, patch),
        },
      };
    }));
    setIsDirty(true);
  }

  function commitDropdownOptions(address = selectedAddress) {
    const options = splitInputOptions(dropdownOptionsDraft);
    updateCell(address, { inputOptions: options });
    setDropdownOptionsDraft(options.join("\n"));
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

      setRedoStack((redoHistory) => [...redoHistory.slice(Math.max(0, redoHistory.length - historyLimit + 1)), makeSnapshot(cells, columnCount, rowCount)]);
      setCells(cloneCells(previous.cells));
      setColumnCount(previous.columnCount);
      setRowCount(previous.rowCount);
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

      setUndoStack((undoHistory) => [...undoHistory.slice(Math.max(0, undoHistory.length - historyLimit + 1)), makeSnapshot(cells, columnCount, rowCount)]);
      setCells(cloneCells(next.cells));
      setColumnCount(next.columnCount);
      setRowCount(next.rowCount);
      setEditingAddress(null);
      setDraftEntry("");
      setIsDirty(true);
      return history.slice(0, -1);
    });
  }

  function startEditing(address: string, replacement?: string) {
    const cell = getCell(cells, address);
    setSelectedAddress(address);
    setEditingAddress(address);
    setDraftEntry(replacement ?? cell.entry);
  }

  function commitEditing(nextAddress?: string) {
    if (!editingAddress) {
      if (nextAddress) setSelectedAddress(nextAddress);
      return;
    }
    updateCell(editingAddress, { entry: draftEntry });
    setEditingAddress(null);
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
    setDraftEntry("");
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
    if (canUseReferencePopup && (event.key === "Tab" || event.key === "Enter")) {
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
          next[targetAddress] = { ...existing, entry };
        });
      });
      return next;
    });
  }

  function resetSheet() {
    applyCellsChange(() => initialCells);
    setColumnCount(defaultColumnCount);
    setRowCount(defaultRowCount);
    setSelectedAddress("B2");
    setEditingAddress(null);
  }

  function clearSheet() {
    applyCellsChange(() => ({}));
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
  }

  function currentWorkbookSheets(): WorkbookSheet[] {
    if (sheets.length === 0) return [makeWorkbookSheet("Sheet 1", cells, columnCount, rowCount)];
    return sheets.map((sheet) => (
      sheet.id === activeSheetId
        ? { ...sheet, cells: hydrateCells(cells), columnCount, rowCount }
        : sheet
    ));
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
        return {
          ...configuration,
          name,
          activeSheetId: nextActiveSheet.id,
          sheets: nextSheets,
          cells: nextActiveSheet.cells,
          columnCount: nextActiveSheet.columnCount,
          rowCount: nextActiveSheet.rowCount,
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
    applySheetResize((current) => insertRow(current, insertAt), columnCount, Math.max(rowCount + 1, insertAt));
    setSelectedAddress(`${selected?.column ?? "A"}${insertAt}`);
    setEditingAddress(null);
  }

  function deleteSelectedRow() {
    const selected = parseAddress(selectedAddress);
    if (!selected || rowCount <= 1) return;
    applySheetResize((current) => deleteRow(current, selected.row), columnCount, Math.max(1, rowCount - 1));
    setSelectedAddress(`${selected.column}${Math.min(selected.row, rowCount - 1)}`);
    setEditingAddress(null);
  }

  function addColumnRight() {
    const selected = parseAddress(selectedAddress);
    const selectedColumnNumber = selected ? columnNumber(selected.column) : columnCount;
    const insertAt = selectedColumnNumber + 1;
    applySheetResize((current) => insertColumn(current, insertAt), Math.max(columnCount + 1, insertAt), rowCount);
    setSelectedAddress(`${columnName(insertAt) ?? "A"}${selected?.row ?? 1}`);
    setEditingAddress(null);
  }

  function deleteSelectedColumn() {
    const selected = parseAddress(selectedAddress);
    if (!selected || columnCount <= 1) return;
    const deleteAt = columnNumber(selected.column);
    applySheetResize((current) => deleteColumn(current, deleteAt), Math.max(1, columnCount - 1), rowCount);
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

    if (configurations.length <= 1) {
      const created = makeConfiguration("Untitled Configuration", {}, defaultColumnCount, defaultRowCount);
      setConfigurations([created]);
      loadConfiguration(created);
      return;
    }

    if (!window.confirm(`Delete "${label}"? This only removes the local browser copy.`)) return;

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
    setSheets(nextSheets);
    setActiveSheetId(nextActiveSheet.id);
    setCells(nextActiveSheet.cells);
    setColumnCount(nextActiveSheet.columnCount);
    setRowCount(nextActiveSheet.rowCount);
    setUndoStack([]);
    setRedoStack([]);
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
    setIsDirty(false);
  }

  function switchSheet(nextSheetId: string) {
    if (nextSheetId === activeSheetId) return;
    const nextSheets = currentWorkbookSheets();
    const nextSheet = nextSheets.find((sheet) => sheet.id === nextSheetId);
    if (!nextSheet) return;

    setSheets(nextSheets);
    setActiveSheetId(nextSheet.id);
    setCells(nextSheet.cells);
    setColumnCount(nextSheet.columnCount);
    setRowCount(nextSheet.rowCount);
    setUndoStack([]);
    setRedoStack([]);
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
    setCopiedAddress(null);
    setIsDirty(true);
  }

  function addSheet() {
    const nextSheets = currentWorkbookSheets();
    const created = makeWorkbookSheet(`Sheet ${nextSheets.length + 1}`, {}, defaultColumnCount, defaultRowCount);
    setSheets([...nextSheets, created]);
    setActiveSheetId(created.id);
    setCells(created.cells);
    setColumnCount(created.columnCount);
    setRowCount(created.rowCount);
    setUndoStack([]);
    setRedoStack([]);
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
    setIsDirty(true);
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
    setSheets(nextSheets);
    if (nextActiveSheet) {
      setCells(nextActiveSheet.cells);
      setColumnCount(nextActiveSheet.columnCount);
      setRowCount(nextActiveSheet.rowCount);
    }
    setIsDirty(true);
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
      const converted = convertImportedSheetToQuoin(sheet, { names: pendingImport.names });
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

  const visibleSheets = sheets.length > 0
    ? sheets.map((sheet) => (
      sheet.id === activeSheetId
        ? { ...sheet, cells, columnCount, rowCount }
        : sheet
    ))
    : [];
  const workbookEngineSheets = useMemo(
    () => visibleSheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      cells: toEngineCells(sheet.cells),
    })),
    [visibleSheets],
  );
  const workbookResult = useMemo(() => executeWorkbookEngine({ sheets: workbookEngineSheets }), [workbookEngineSheets]);
  const runnerSheets = useMemo(
    () => buildRunnerSheetContexts(visibleSheets, workbookResult),
    [visibleSheets, workbookResult],
  );
  const selectedIssues = issueMap.get(selectedAddress) ?? [];
  const selectedImportSheet = pendingImport?.sheets.find((sheet) => sheet.id === selectedImportSheetId) ?? pendingImport?.sheets[0] ?? null;
  const importReviewItems = pendingImport && selectedImportSheet
    ? pendingImport.reviewItems.concat(importReviewItemsForSheet(pendingImport.names, selectedImportSheet.name))
    : [];

  return (
    <>
      <header className="topbar">
        <div>
          <p className="eyebrow">Quoin Core</p>
          <h1>Variable Sheet</h1>
        </div>
        <div className="toolbar">
          <div className="configBar" aria-label="Local configurations">
            <select
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
          <button type="button" onClick={createConfiguration}>New</button>
          <button type="button" onClick={saveConfiguration}>Save</button>
          <button type="button" onClick={duplicateConfiguration}>Duplicate</button>
          <button type="button" onClick={deleteConfiguration}>Delete</button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            className="fileInput"
            onChange={handleImportFileChange}
          />
          <button type="button" onClick={() => fileInputRef.current?.click()} disabled={isImporting}>
            {isImporting ? "Importing..." : "Import Excel"}
          </button>
          <button type="button" onClick={undoCells} disabled={undoStack.length === 0}>Undo</button>
          <button type="button" onClick={redoCells} disabled={redoStack.length === 0}>Redo</button>
          <button type="button" onClick={() => copyCell(selectedAddress)}>Copy Cell</button>
          <button type="button" onClick={() => pasteCopiedCell(selectedAddress)} disabled={!copiedAddress}>Paste Cell</button>
          <button type="button" onClick={() => fillDown(selectedAddress)}>Fill Down</button>
          <button type="button" onClick={addRowBelow}>Add Row</button>
          <button type="button" onClick={deleteSelectedRow}>Delete Row</button>
          <button type="button" onClick={addColumnRight}>Add Column</button>
          <button type="button" onClick={deleteSelectedColumn}>Delete Column</button>
          <button type="button" onClick={clearSheet}>Clear Sheet</button>
          <button type="button" onClick={resetSheet}>Load Demo</button>
          <div className="status" data-valid={result.valid}>
            {result.valid ? "Engine Ready" : "Engine Error"}
          </div>
        </div>
      </header>

      {(pendingImport || importError) && (
        <section className="importPanel" aria-label="Excel import">
          <div className="importPanelHeader">
            <div>
              <p className="eyebrow">Excel Import</p>
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

      <div className="viewTabs" role="tablist" aria-label="Quoin surfaces">
        <button type="button" data-active={activeView === "sheet"} onClick={() => setActiveView("sheet")}>
          Sheet
        </button>
        <button type="button" data-active={activeView === "runner"} onClick={() => setActiveView("runner")}>
          Runner Preview
        </button>
        <button type="button" data-active={activeView === "help"} onClick={() => setActiveView("help")}>
          Help
        </button>
      </div>

      {activeView === "sheet" ? (
        <>
          <div className="formulaBar">
            <span>{activeSheet ? `${activeSheet.name}!${selectedAddress}` : selectedAddress}</span>
            <input
              aria-label="Formula bar"
              value={editingAddress === selectedAddress ? draftEntry : selectedCell.entry}
              onChange={(event) => {
                if (editingAddress !== selectedAddress) setEditingAddress(selectedAddress);
                setDraftEntry(event.target.value);
              }}
              onKeyDown={handleFormulaBarKeyDown}
              onBlur={() => editingAddress === selectedAddress && commitEditing()}
              onFocus={() => {
                setEditingAddress(selectedAddress);
                setDraftEntry(selectedCell.entry);
              }}
              placeholder="Value or formula"
            />
            {editingAddress && draftEntry.trim().startsWith("=") && (
              <div className="referencePopup">
                <div className="referencePopupHeader">Insert reference</div>
                {visibleReferenceOptions
                  .map((option, index) => (
                    <button
                      key={option.address}
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

          <div className="authoringLayout">
            <section className="spreadsheetFrame" aria-label="Quoin spreadsheet grid">
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

                {Array.from({ length: rowCount }, (_, rowIndex) => {
                  const rowNumber = rowIndex + 1;
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
              </div>
            </section>

            <Inspector
              clearCell={clearCell}
              dependencySummary={dependencySummary}
              displayValue={displayValues[selectedAddress] ?? ""}
              dropdownOptionsDraft={dropdownOptionsDraft}
              setDropdownOptionsDraft={setDropdownOptionsDraft}
              commitDropdownOptions={commitDropdownOptions}
              selectedAddress={selectedAddress}
              selectedCell={selectedCell}
              selectedIssues={selectedIssues}
              updateCell={updateCell}
              updateLookup={updateLookup}
            />
          </div>
        </>
      ) : activeView === "runner" ? (
        <RunnerPreview
          result={workbookResult}
          runnerSheets={runnerSheets}
          updateCell={updateWorkbookCell}
        />
      ) : (
        <HelpPanel />
      )}

      {(Object.keys(workbookResult.outputs).length > 0 || workbookResult.warnings.length > 0) && (
        <section className={`runnerStrip ${workbookResult.warnings.length === 0 ? "runnerStripSingle" : ""}`}>
          {Object.keys(workbookResult.outputs).length > 0 && (
            <div>
              <span>Surfaced Results</span>
              <strong>{formatOutputs(workbookResult.outputs)}</strong>
            </div>
          )}
          {workbookResult.warnings.length > 0 && (
            <div data-kind="warning">
              <span>Review Needed</span>
              <strong>{formatWorkbookWarnings(workbookResult.warnings, runnerSheets)}</strong>
            </div>
          )}
        </section>
      )}
    </>
  );
}

function SheetStrip({
  activeSheetId,
  addSheet,
  renameSheet,
  sheets,
  switchSheet,
}: {
  activeSheetId: string;
  addSheet: () => void;
  renameSheet: (sheetId: string, name: string) => void;
  sheets: WorkbookSheet[];
  switchSheet: (sheetId: string) => void;
}) {
  const activeSheet = sheets.find((sheet) => sheet.id === activeSheetId) ?? sheets[0];

  return (
    <aside className="sheetStrip" aria-label="Sheets">
      <div className="sheetStripHeader">Sheets</div>
      <div className="sheetStripList" role="tablist" aria-label="Workbook sheets">
        {sheets.map((sheet) => (
          <button
            key={sheet.id}
            type="button"
            data-active={sheet.id === activeSheetId}
            onClick={() => switchSheet(sheet.id)}
            role="tab"
            aria-selected={sheet.id === activeSheetId}
            title={sheet.name}
          >
            <span>{sheet.name}</span>
            <small>{Object.keys(sheet.cells).length} cells</small>
          </button>
        ))}
      </div>
      <button type="button" className="sheetAddButton" onClick={addSheet}>
        Add Sheet
      </button>
      {activeSheet && (
        <label className="sheetRename">
          <span>Active Sheet</span>
          <input
            value={activeSheet.name}
            onChange={(event) => renameSheet(activeSheet.id, event.target.value)}
          />
        </label>
      )}
    </aside>
  );
}

function Inspector({
  clearCell,
  commitDropdownOptions,
  dependencySummary,
  displayValue,
  dropdownOptionsDraft,
  selectedAddress,
  selectedCell,
  selectedIssues,
  setDropdownOptionsDraft,
  updateCell,
  updateLookup,
}: {
  clearCell: (address: string) => void;
  commitDropdownOptions: (address?: string) => void;
  dependencySummary: DependencySummary;
  displayValue: CellValue;
  dropdownOptionsDraft: string;
  selectedAddress: string;
  selectedCell: GridCell;
  selectedIssues: string[];
  setDropdownOptionsDraft: (value: string) => void;
  updateCell: (address: string, patch: Partial<GridCell>) => void;
  updateLookup: (patch: Partial<LookupConfig>) => void;
}) {
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
          <button type="button" onClick={() => clearCell(selectedAddress)}>Clear Cell</button>
        </div>
      </div>

      {!selectedCell.name && (
        <div className="activationBox">
          <strong>Normal spreadsheet cell</strong>
          <p>Add a Smart Cell name below when this cell needs metadata, runner surfacing, or named formula references.</p>
        </div>
      )}

      <div className="inspectorSection">
        <div className="sectionTitle">
          <strong>Cell</strong>
          <span>{selectedCell.name ? "Spreadsheet value and Smart Cell identity" : "Spreadsheet value first"}</span>
        </div>

        <label>
          Cell Entry
          <input value={selectedCell.entry} onChange={(event) => updateCell(selectedAddress, { entry: event.target.value })} />
        </label>

        <label>
          Smart Cell Name
          <input
            placeholder="example: wall_height"
            value={selectedCell.name}
            onChange={(event) => {
              const name = sanitizeName(event.target.value);
              updateCell(selectedAddress, {
                name,
                surfaced: name ? selectedCell.surfaced : false,
              });
            }}
          />
        </label>

        {selectedCell.name && (
          <label>
            Display Label
            <input
              placeholder={prettifyName(selectedCell.name)}
              value={selectedCell.label}
              onChange={(event) => updateCell(selectedAddress, { label: event.target.value })}
            />
          </label>
        )}
      </div>

      {selectedCell.name && (
        <div className="inspectorSection smartSection">
          <div className="sectionTitle">
            <strong>Smart Behavior</strong>
            <span>Metadata for runner preview and named calculations</span>
          </div>
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
                {roleOptions.map((role) => (
                  <option key={role} value={role}>{role}</option>
                ))}
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
            {selectedCell.role === "input" && (
              <label>
                Input Control
                <select
                  value={selectedCell.inputControl}
                  onChange={(event) => {
                    const inputControl = event.target.value as InputControl;
                    const options = inputControl === "dropdown" && selectedCell.inputOptions.length === 0 && selectedCell.entry.trim()
                      ? [selectedCell.entry.trim()]
                      : selectedCell.inputOptions;
                    setDropdownOptionsDraft(options.join("\n"));
                    updateCell(selectedAddress, { inputControl, inputOptions: inputControl === "dropdown" ? options : [] });
                  }}
                >
                  {inputControlOptions.map((inputControl) => (
                    <option key={inputControl} value={inputControl}>{inputControl === "freeText" ? "Free text" : "Dropdown"}</option>
                  ))}
                </select>
              </label>
            )}
          </div>

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

          <label>
            Internal Annotation
            <textarea
              value={selectedCell.annotation}
              onChange={(event) => updateCell(selectedAddress, { annotation: event.target.value })}
              placeholder="Internal note about what this cell means"
              rows={4}
            />
          </label>

          {selectedCell.role === "input" && selectedCell.inputControl === "dropdown" && (
            <label>
              Dropdown Options
              <textarea
                value={dropdownOptionsDraft}
                onBlur={() => commitDropdownOptions(selectedAddress)}
                onChange={(event) => setDropdownOptionsDraft(event.target.value)}
                placeholder="One short option per line, or comma-separated. Leave blank for free text."
                rows={3}
              />
              <span>Use embedded options for short lists. Longer lists should come from visible reference data in a later reference-table workflow.</span>
            </label>
          )}

          {(selectedCell.role === "validation" || selectedCell.role === "compliance") && (
            <label>
              Runner Message
              <textarea
                value={selectedCell.ruleMessage}
                onChange={(event) => updateCell(selectedAddress, { ruleMessage: event.target.value })}
                placeholder={selectedCell.role === "validation" ? "Example: Input is outside the approved range." : "Example: Manual review is recommended."}
                rows={3}
              />
            </label>
          )}
        </div>
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

      {selectedIssues.length > 0 && (
        <div className="issueBox">
          {selectedIssues.map((issue) => (
            <p key={issue}>{issue}</p>
          ))}
        </div>
      )}
    </aside>
  );
}

function RunnerPreview({
  result,
  updateCell,
  runnerSheets,
}: {
  result: WorkbookEngineResult;
  runnerSheets: RunnerSheetContext[];
  updateCell: (sheetId: string, address: string, patch: Partial<GridCell>) => void;
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

  return (
    <section className="runnerPreview">
      <div className="runnerHeader">
        <div>
          <p className="eyebrow">Runner Preview</p>
          <h2>Generated Form</h2>
        </div>
        <span data-valid={result.valid}>{result.valid ? "Ready" : "Failed Validation"}</span>
      </div>

      <div className={`runnerGrid ${outputGroups.length === 0 ? "runnerGridSingle" : ""}`}>
        <div className="runnerPanel">
          <h3>Inputs</h3>
          {inputGroups.length === 0 ? (
            <p className="runnerEmpty">No surfaced inputs.</p>
          ) : (
            inputGroups.map((group) => (
              <RunnerSheetGroup key={group.sheetId} showHeading={showSheetGroups} sheetName={group.sheetName}>
                {group.items.map((cell) => (
                  <label key={`${group.sheetId}-${cell.address}`}>
                    {labelForCell(cell)}
                    {isDropdownCell(cell) ? (
                      <select value={cell.entry} onChange={(event) => updateCell(group.sheetId, cell.address, { entry: event.target.value })}>
                        {!cell.entry && <option value="">Choose...</option>}
                        {dropdownOptionsForCell(cell).map((option) => (
                          <option key={option} value={option}>{prettifyName(option)}</option>
                        ))}
                      </select>
                    ) : (
                      <input value={cell.entry} onChange={(event) => updateCell(group.sheetId, cell.address, { entry: event.target.value })} />
                    )}
                    {cell.annotation && <small>{cell.annotation}</small>}
                  </label>
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
                {group.items.map((cell) => (
                  <div className="runnerResult" key={`${group.sheetId}-${cell.address}`}>
                    <span>{labelForCell(cell)}</span>
                    <strong>{formatCellValue(group.displayValues[cell.address] ?? null)}</strong>
                    {cell.annotation && <small>{cell.annotation}</small>}
                  </div>
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
                  {group.items.map((cell) => (
                    <p data-state="action" key={`${group.sheetId}-${cell.address}`}>
                      <strong>ACTION</strong>
                      {formatCellValue(group.displayValues[cell.address] ?? null) || labelForCell(cell)}
                    </p>
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
                  {group.items.map((warning) => (
                    <p data-state="warn" key={warning.cellId}>
                      <strong>WARN</strong>
                      {warning.message}
                    </p>
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
                {group.items.map((rule) => {
                  const cell = getCell(group.cells, rule.address);
                  return (
                    <p data-state={rule.state} key={`${group.sheetId}-${rule.address}`}>
                      <strong>{formatCellValue(group.displayValues[rule.address] ?? null)}</strong>
                      {cell.ruleMessage || cell.annotation || labelForCell(cell)}
                    </p>
                  );
                })}
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

function HelpPanel() {
  return (
    <section className="helpPanel" aria-label="Quoin help">
      <div className="helpHeader">
        <p className="eyebrow">Help</p>
        <h2>How Quoin Works</h2>
        <p>Quoin is a spreadsheet-first prototype for turning shop knowledge into structured, runner-safe workflows. Start by building normal spreadsheet logic, then name important cells when they need meaning, controls, or runner visibility.</p>
      </div>

      <div className="helpGrid">
        <article>
          <h3>Core Idea</h3>
          <p>Quoin should feel familiar to anyone who has used Excel: cells have coordinates, formulas can reference other cells, and imported workbooks keep their Sheet structure. The difference is that named cells can carry metadata, controls, runner labels, validation, and review behavior.</p>
          <p>The authoring grid is still the source of truth. Build and import calculator logic first, prove the math, then promote the important cells into Smart Cells for runner-facing workflow behavior.</p>
        </article>

        <article>
          <h3>Normal Cells</h3>
          <p>Normal cells use addresses like <code>A1</code> or <code>B12</code>. They can hold labels, numbers, text, booleans, or formulas. They stay in the authoring grid only and do not appear in Runner Preview.</p>
          <ul>
            <li>Type directly in a selected cell, double-click a cell, press Enter, or use the formula bar.</li>
            <li>Use Delete or Backspace to clear the selected cell.</li>
            <li>Paste tabular data from Excel into the grid.</li>
          </ul>
        </article>

        <article>
          <h3>Smart Cells</h3>
          <p>A cell becomes a Smart Cell when you give it a Smart Cell Name in the inspector. Names are formula-safe identifiers such as <code>design_span</code>, while Display Label is the human-facing text such as <code>Design Span (ft)</code>.</p>
          <ul>
            <li>Smart Cell names are workbook-scoped when unique.</li>
            <li>Formulas on any Sheet can reference a unique Smart Cell name directly.</li>
            <li>Metadata stays in the inspector so the grid remains spreadsheet-first.</li>
            <li>Naming a cell does not automatically show it to the runner. Turn on Surface to runner when the cell belongs in Runner Preview.</li>
            <li>Imported Excel dropdown cells are the exception: supported dropdown inputs are surfaced automatically so the imported control remains visible where it is used.</li>
          </ul>
        </article>

        <article>
          <h3>Basic Workflow</h3>
          <ol>
            <li>Build the calculator in the grid with normal values and formulas.</li>
            <li>Verify the math works in the Sheet view.</li>
            <li>Name important cells to promote them to Smart Cells.</li>
            <li>Set role, value type, input control, display label, annotation, dropdown options, or rule text in the inspector.</li>
            <li>Turn on Surface to runner only for the cells the runner should see.</li>
            <li>Use Runner Preview to test the controlled form.</li>
          </ol>
        </article>

        <article>
          <h3>Roles</h3>
          <p>Roles describe how a Smart Cell behaves and how it should be grouped for the runner.</p>
          <ul>
            <li><code>input</code>: a value supplied by an admin or runner.</li>
            <li><code>formula</code>: internal calculated logic.</li>
            <li><code>output</code>: a calculated or entered result.</li>
            <li><code>lookup</code>: a prototype value resolved from an embedded lookup table.</li>
            <li><code>action</code>: a runner-facing shop note or required action.</li>
            <li><code>validation</code>: PASS or FAIL rule result.</li>
            <li><code>compliance</code>: OK or WARN review result.</li>
          </ul>
          <p>The current lookup role is useful for small demos, but imported shop calculators point toward first-class Reference Tables instead of large hidden lookup tables inside one Smart Cell.</p>
        </article>

        <article>
          <h3>Inputs And Dropdowns</h3>
          <p>Input Smart Cells can use a free text control or a controlled dropdown control. Choose Dropdown in Input Control, then add short embedded choices in the inspector with one option per line or comma-separated values.</p>
          <ul>
            <li>Dropdowns render directly in the grid.</li>
            <li>The same dropdown options render in Runner Preview.</li>
            <li>Use Display Label to make runner fields readable without changing formula names.</li>
            <li>Imported Excel cells with simple typed lists or same-workbook range lists are snapshotted into embedded dropdown options.</li>
            <li>Longer option sets should eventually use live visible reference data, named ranges, tables, or CSV-ingested datasets.</li>
          </ul>
        </article>

        <article className="helpFormulaReference">
          <h3>Supported Formulas</h3>
          <p>Quoin supports an Excel-compatible subset for common calculator logic. A formula starts with <code>=</code>. The engine evaluates references, ranges, arithmetic, comparisons, common functions, and <code>IF</code>. Unsupported imported formulas stay visible and appear as review items instead of being silently dropped.</p>
          <h4>References</h4>
          <p>A reference points a formula at another cell or Smart Cell. Coordinate references use the grid address. Smart Cell references use the workbook-scoped Smart Cell Name.</p>
          <ul>
            <li><code>=A1+B1</code> uses coordinate references.</li>
            <li><code>=B2 * 650</code> multiplies the value in <code>B2</code> by a constant.</li>
            <li><code>=design_span * design_plf</code> uses Smart Cell names.</li>
            <li><code>=design_span * 650</code> can reference <code>design_span</code> from another Sheet.</li>
          </ul>
          <h4>Cross-Sheet References</h4>
          <p>Use cross-Sheet references when a formula should point at a coordinate on another Sheet. Sheet names with spaces need single quotes.</p>
          <ul>
            <li><code>=Inputs!B2 * Inputs!B3</code> references cells on a Sheet named <code>Inputs</code>.</li>
            <li><code>='Input Data'!B2 * 3</code> references a Sheet whose name contains a space.</li>
            <li><code>=design_span * design_plf</code> is usually cleaner when those inputs have unique Smart Cell names.</li>
          </ul>
          <h4>Ranges</h4>
          <p>A range is a group of cells between two addresses. Quoin supports single-column, single-row, and rectangular ranges for common aggregate formulas.</p>
          <ul>
            <li><code>=SUM(A1:A5)</code> and <code>=SUM(A1:B3)</code> use ranges.</li>
            <li><code>=AVERAGE(B2:B10)</code> averages a vertical range.</li>
            <li><code>=MAX(A1:D1)</code> finds the largest value across a row.</li>
            <li><code>=SUM(Loads!B3:B5)</code> can aggregate a supported range from another Sheet.</li>
          </ul>
          <h4>Functions</h4>
          <p>Functions perform named operations. Quoin accepts familiar uppercase Excel-style names and maps them to deterministic engine behavior.</p>
          <ul>
            <li><code>=SUM(A1:A5)</code> adds values.</li>
            <li><code>=AVERAGE(A1:A5)</code> calculates the mean.</li>
            <li><code>=MIN(A1:A5)</code> and <code>=MAX(A1:A5)</code> find bounds.</li>
            <li><code>=ROUND(B6, 2)</code> rounds to two decimal places.</li>
            <li><code>=ROUNDUP(B6, 0)</code> rounds away from zero, matching the common Excel calculator pattern.</li>
            <li><code>=ABS(B2)</code>, <code>=SQRT(B2)</code>, <code>=CEIL(B2)</code>, and <code>=FLOOR(B2)</code> cover common numeric cleanup.</li>
          </ul>
          <h4>Conditions And IF</h4>
          <p>Comparisons return true or false. <code>IF</code> chooses one value when the condition is true and another value when it is false.</p>
          <ul>
            <li><code>=design_span &gt; 14</code> returns a boolean result.</li>
            <li><code>=IF(A1&gt;10, "review", "ok")</code> uses conditional logic.</li>
            <li><code>=IF(design_span&gt;14, "review", recommended_beam)</code> combines Smart Cell names and IF.</li>
            <li><code>=IF(total_line_load&gt;9000, "engineering review", "standard")</code> returns runner-readable text.</li>
          </ul>
          <p>Supported aliases include SUM, AVERAGE, MAX, MIN, ROUND, ROUNDUP, ABS, SQRT, CEIL, and FLOOR.</p>
          <h4>Formula Editing Notes</h4>
          <ul>
            <li>While typing a formula, the reference popup suggests named Smart Cells and populated coordinates.</li>
            <li>Copy, paste, and fill-down adjust coordinate references while leaving Smart Cell names unchanged.</li>
            <li>Deleted row or column references become <code>#REF!</code> so broken formulas remain visible.</li>
            <li>Renaming a Sheet updates direct cross-Sheet references that use the old Sheet name.</li>
            <li>Incomplete formulas should not crash the app while you are still typing.</li>
          </ul>
        </article>

        <article>
          <h3>Sheets</h3>
          <p>A configuration can contain multiple Sheets. Sheet tabs sit below the formula bar and above the grid. The authoring grid shows one active Sheet at a time, while Runner Preview can gather surfaced Smart Cells from the whole workbook.</p>
          <ul>
            <li>Use the Active Sheet field in the Sheet strip to rename the current Sheet.</li>
            <li>Use coordinate formulas on the current Sheet, such as <code>=A1+B1</code>.</li>
            <li>Use supported cross-Sheet references, such as <code>=Inputs!B2 * Inputs!B3</code>.</li>
            <li>Use unique Smart Cell names across Sheets, such as <code>=design_span * design_plf</code>.</li>
            <li>Column headers stay visible while scrolling down the grid, and row numbers stay visible while scrolling sideways.</li>
          </ul>
        </article>

        <article>
          <h3>Runner Preview</h3>
          <p>Runner Preview is generated from surfaced Smart Cells. It hides normal coordinate cells and empty optional sections, then groups surfaced inputs, outputs, actions, review flags, and validation by Sheet when needed.</p>
          <ul>
            <li>Inputs become editable runner fields.</li>
            <li>Outputs show calculated results.</li>
            <li>Action cells show shop notes or required actions.</li>
            <li>Validation cells show PASS or FAIL.</li>
            <li>Compliance cells show OK or WARN.</li>
            <li>Empty output, action, review, and validation sections stay hidden so the runner view stays focused.</li>
            <li>The bottom summary appears only when there are surfaced results or review warnings.</li>
          </ul>
        </article>

        <article>
          <h3>Validation vs Compliance</h3>
          <p>Validation and compliance are intentionally different. Validation is for run failure; compliance is for warning the runner. Math should still run where possible.</p>
          <ul>
            <li>Validation true means PASS. Validation false means FAIL.</li>
            <li>Compliance false means OK. Compliance true means WARN.</li>
            <li>Rule Message is the runner-facing explanation for a FAIL or WARN.</li>
          </ul>
        </article>

        <article>
          <h3>Lookup Cells</h3>
          <p>Lookup Smart Cells can match one or more input criteria against the embedded lookup table editor. The current editor is a prototype shortcut for small examples; the longer-term product direction is first-class Reference Tables or CSV-ingested datasets.</p>
          <ul>
            <li>Use criteria columns to match input Smart Cell names.</li>
            <li>Choose an output column to return the lookup result.</li>
            <li>Paste tabular data from Excel into the lookup table editor.</li>
            <li>A lookup miss shows <code>#ERR</code> so missing data is visible.</li>
            <li>Do not use the embedded lookup editor for thousands of rows. Large imported data Sheets should become Reference Tables in a later workflow.</li>
          </ul>
        </article>

        <article>
          <h3>Reference Data Direction</h3>
          <p>Real shop calculators often use large data tabs and Excel lookup formulas. Quoin now calls out likely reference-data Sheets during import so the admin can tell the difference between a calculator Sheet and a data Sheet.</p>
          <ul>
            <li>A large Sheet with many rows and no formulas is preserved as a normal Sheet for now and marked as likely reference data.</li>
            <li>Exact-match Excel <code>VLOOKUP</code> formulas are preserved, but review items explain the source range, lookup key, output column, and Reference Table repair path.</li>
            <li>Quoin should eventually bind these formulas to visible/imported Reference Tables instead of hiding thousands of rows inside Smart Cell metadata.</li>
            <li>Runner Preview should show surfaced inputs and results, not raw reference-data rows.</li>
          </ul>
        </article>

        <article>
          <h3>Excel Import</h3>
          <p>Import Excel brings an <code>.xlsx</code> calculator into a new browser-local configuration. The goal is to preserve the workbook calculation surface first, then let you structure it with Quoin Smart Cells.</p>
          <ul>
            <li>Workbook Sheets, values, and formulas are preserved.</li>
            <li>Safe workbook-defined single-cell names can become Smart Cell names.</li>
            <li>Simple Excel data-validation lists and supported workbook range sources can become dropdown options on imported input Smart Cells.</li>
            <li>Dropdown-only cells still expand the imported Sheet bounds, so controls outside the normal used range remain visible.</li>
            <li>Merged ranges, named ranges, external workbook links, structured table references, spill markers, and other risky features appear as review items.</li>
            <li>Unsupported formulas remain visible instead of being silently dropped.</li>
            <li>Review items try to explain the repair path, such as converting an exact-match <code>VLOOKUP</code> into a Reference Table lookup or replacing <code>INDIRECT</code> with direct references.</li>
          </ul>
        </article>

        <article>
          <h3>Formula Review Items</h3>
          <p>Imported Excel formulas are classified conservatively. Quoin supports the common subset it can evaluate deterministically and preserves the rest for review.</p>
          <ul>
            <li><code>VLOOKUP</code>: exact-match lookups should become Reference Table lookups. Approximate or omitted match-mode lookups need manual confirmation.</li>
            <li><code>IFERROR</code> and <code>IFNA</code>: decide what fallback is safe, then model it with explicit <code>IF</code>, validation, or review messaging.</li>
            <li><code>INDIRECT</code> and <code>OFFSET</code>: replace dynamic address logic with direct references or a Reference Table selection.</li>
            <li><code>SUMIFS</code>, <code>COUNTIFS</code>, and similar criteria formulas: move the criteria ranges into a Reference Table or helper calculation before modeling the aggregate.</li>
            <li>Date and time functions need a separate date-semantics decision before Quoin should evaluate them.</li>
          </ul>
        </article>

        <article>
          <h3>Local Configurations</h3>
          <p>Configurations are stored in this browser for now. New, Save, Duplicate, Delete, rename, Load Demo, and Import Excel all operate locally. Database-backed publishing, execution records, auth, and audit reports are later phases.</p>
        </article>

        <article>
          <h3>Beam Demo</h3>
          <p>The default <code>Demo - Beam Selection</code> configuration uses fake data only. It shows the intended flow: runner inputs for drawing conditions, calculated context, lookup recommendations, shop action notes, validation, and compliance warnings.</p>
        </article>

        <article>
          <h3>Keyboard Shortcuts</h3>
          <ul>
            <li>Arrow keys move the selected cell.</li>
            <li>Enter or F2 edits the selected cell.</li>
            <li>Tab moves right; Shift+Tab moves left.</li>
            <li>Ctrl+Z and Ctrl+Y undo and redo sheet edits.</li>
            <li>Ctrl+C, Ctrl+V, and Ctrl+D support cell copy, paste, and fill-down behavior.</li>
          </ul>
        </article>
      </div>
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
        const dropdownOptions = dropdownOptionsForCell(cell);
        return (
          <div
            className="gridCell"
            data-dropdown={hasDropdown}
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
        lookup: cell.lookup ?? (cell.role === "lookup" || cell.role === "action" ? starterLookup : undefined),
      };

      return [address, applyCellPatch(hydrated, {})];
    }),
  );
}

function cloneCells(cells: Record<string, GridCell>): Record<string, GridCell> {
  return JSON.parse(JSON.stringify(cells)) as Record<string, GridCell>;
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
  const rawExpression = entry.trim().startsWith("=") ? entry.trim().slice(1) : entry;
  const expression = rawExpression.replace(/"[^"]*"|'[^']*'/g, " ");
  const refs = new Set<string>();

  for (const range of expression.matchAll(/\b([A-Z]+[1-9]\d*)\s*:\s*([A-Z]+[1-9]\d*)\b/g)) {
    const expanded = expandAddressRange(range[1], range[2]);
    for (const address of expanded) refs.add(address);
  }

  for (const match of expression.replace(/\b([A-Z]+[1-9]\d*)\s*:\s*([A-Z]+[1-9]\d*)\b/g, " ").matchAll(/\b[A-Za-z_][A-Za-z0-9_]*\b/g)) {
    const token = match[0];
    if (isFormulaKeyword(token)) continue;
    refs.add(token);
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

function isFormulaKeyword(token: string): boolean {
  return new Set([
    "SUM",
    "sum",
    "AVERAGE",
    "average",
    "mean",
    "MAX",
    "max",
    "MIN",
    "min",
    "ROUND",
    "round",
    "ABS",
    "abs",
    "SQRT",
    "sqrt",
    "CEIL",
    "ceil",
    "FLOOR",
    "floor",
    "IF",
    "if",
    "true",
    "false",
    "LOOKUP",
  ]).has(token);
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

function makeSnapshot(cells: Record<string, GridCell>, columnCount: number, rowCount: number): SheetSnapshot {
  return {
    cells: cloneCells(cells),
    columnCount,
    rowCount,
  };
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
  const nextPrefix = sheetReferencePrefix(newName);
  const escapedQuotedName = escapeRegExp(oldName.replace(/'/g, "''"));
  const escapedPlainName = escapeRegExp(oldName);

  return entry
    .replace(new RegExp(`'${escapedQuotedName}'!`, "g"), `${nextPrefix}!`)
    .replace(new RegExp(`\\b${escapedPlainName}!`, "g"), `${nextPrefix}!`);
}

function sheetReferencePrefix(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName)
    ? sheetName
    : `'${sheetName.replace(/'/g, "''")}'`;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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
  return entry.replace(/\b([A-Z]+)([1-9]\d*)\b/g, (_match, column: string, row: string) => replacer(column, Number(row)));
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
        cells: activeSheet.cells,
        columnCount: activeSheet.columnCount,
        rowCount: activeSheet.rowCount,
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
  const activeSheet = workbook?.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook?.sheets[0];
  return {
    id: makeConfigId(),
    name,
    activeSheetId: activeSheet?.id,
    sheets: workbook?.sheets.map(hydrateWorkbookSheet),
    cells: hydrateCells(activeSheet?.cells ?? cells),
    columnCount: Math.max(activeSheet?.columnCount ?? columnCount, defaultColumnCount),
    rowCount: Math.max(activeSheet?.rowCount ?? rowCount, defaultRowCount),
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

  if (hasName && hasFormulaEntry && next.role === "input" && patch.role === undefined) {
    next.role = "formula";
  }

  return next;
}

function isDropdownCell(cell: GridCell): boolean {
  return cell.role === "input" && cell.inputControl === "dropdown";
}

function dropdownOptionsForCell(cell: GridCell): string[] {
  if (!isDropdownCell(cell)) return [];
  if (cell.entry && !cell.inputOptions.includes(cell.entry)) return [cell.entry, ...cell.inputOptions];
  return cell.inputOptions;
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
  return sheets.map((sheet) => {
    const sheetResult = workbookResult.sheetResults.find((item) => item.sheetId === sheet.id);
    const resultForSheet = sheetResult?.result ?? executeEngine({ cells: toEngineCells(sheet.cells) });
    const ruleStateMap = new Map(resultForSheet.ruleStates.map((rule) => [rule.address, rule.state]));
    const columns = makeColumns(sheet.columnCount);
    const displayValues = buildDisplayValues(sheet.cells, resultForSheet.values, resultForSheet.errors, ruleStateMap, columns, sheet.rowCount);
    const surfacedCells = Object.values(sheet.cells).filter((cell) => cell.name && cell.surfaced);
    const validationStates = resultForSheet.ruleStates.filter((rule) => {
      const cell = getCell(sheet.cells, rule.address);
      return cell.role === "validation" && cell.surfaced;
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

function formatWorkbookWarnings(warnings: Array<{ cellId: string; message: string }>, runnerSheets: RunnerSheetContext[]): string {
  const sheetByCellId = new Map<string, string>();

  for (const sheet of runnerSheets) {
    for (const cell of Object.values(sheet.cells)) {
      sheetByCellId.set(`${sheet.sheetId}!${cell.address}`, sheet.sheetName);
    }
  }

  return warnings.map((warning) => {
    const sheetName = sheetByCellId.get(warning.cellId);
    return sheetName ? `${sheetName}: ${warning.message}` : warning.message;
  }).join(" ");
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
        display[address] = values[cell.name || address] ?? (errorAddresses.has(address) ? "#ERR" : "");
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

function buildReferenceOptions(cells: Record<string, GridCell>, displayValues: Record<string, CellValue>, columns: string[], rowCount: number) {
  const options: Array<{ address: string; reference: string; value: CellValue }> = [];

  for (let row = 1; row <= rowCount; row++) {
    for (const column of columns) {
      const address = `${column}${row}`;
      const cell = getCell(cells, address);
      if (!cell.entry && !cell.name) continue;
      options.push({
        address,
        reference: cell.name || address,
        value: displayValues[address] ?? "",
      });
    }
  }

  return options;
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
  return entry.replace(/\b([A-Z]+)([1-9]\d*)\b/g, (match, column: string, row: string) => {
    const nextColumn = columnName(columnNumber(column) + columnOffset);
    const nextRow = Number(row) + rowOffset;
    if (!nextColumn || nextRow < 1) return match;
    return `${nextColumn}${nextRow}`;
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

function formatOutputs(outputs: Record<string, CellValue>): string {
  const entries = Object.entries(outputs);
  if (entries.length === 0) return "No surfaced outputs";
  return entries.map(([key, value]) => `${key}: ${formatCellValue(value)}`).join(" | ");
}
