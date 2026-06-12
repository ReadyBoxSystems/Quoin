"use client";

import { useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import { executeEngine, type CellValue, type EngineCell, type SmartCellRole, type SmartCellType } from "@/lib/engine";
import { convertImportedSheetToQuoin } from "@/lib/import/convert";
import type { ImportedName, ImportedWorkbook, ImportReviewItem } from "@/lib/import/types";
import type { GridCell, LocalConfiguration, LookupConfig, SheetSnapshot } from "@/lib/sheet/types";

const STORAGE_KEY = "quoin.gridSheet.v2";
const CONFIG_STORAGE_KEY = "quoin.configurations.v1";
const ACTIVE_CONFIG_KEY = "quoin.activeConfiguration.v1";
const defaultColumnCount = 8;
const defaultRowCount = 18;
const historyLimit = 50;
const roleOptions: SmartCellRole[] = ["input", "formula", "output", "action", "lookup", "validation", "compliance"];
const typeOptions: SmartCellType[] = ["number", "text", "boolean"];

interface DependencyItem {
  address: string;
  label: string;
  reference: string;
}

interface DependencySummary {
  dependencies: DependencyItem[];
  dependents: DependencyItem[];
}

interface ImportReport {
  fileName: string;
  sheetName: string;
  cellCount: number;
  formulaCount: number;
  promotedNameCount: number;
  reviewItems: ImportReviewItem[];
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
  const [lastImportReport, setLastImportReport] = useState<ImportReport | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);
  const cellRefs = useRef<Record<string, HTMLDivElement | null>>({});

  const columns = useMemo(() => makeColumns(columnCount), [columnCount]);
  const selectedCell = getCell(cells, selectedAddress);
  const engineCells = useMemo(() => toEngineCells(cells), [cells]);
  const result = useMemo(() => executeEngine({ cells: engineCells }), [engineCells]);
  const ruleStateMap = useMemo(() => new Map(result.ruleStates.map((rule) => [rule.address, rule.state])), [result.ruleStates]);
  const displayValues = useMemo(
    () => buildDisplayValues(cells, result.values, result.errors, ruleStateMap, columns, rowCount),
    [cells, columns, result.errors, result.values, rowCount, ruleStateMap],
  );
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
      setCells(activeConfig.cells);
      setColumnCount(activeConfig.columnCount ?? defaultColumnCount);
      setRowCount(activeConfig.rowCount ?? defaultRowCount);
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
    if (!isLoaded) return;

    const smartCells = Object.values(cells).filter((cell) => cell.name);
    if (smartCells.length === 0 || smartCells.some((cell) => cell.surfaced)) return;

    setCells((current) => {
      return Object.fromEntries(
        Object.entries(current).map(([address, cell]) => [
          address,
          cell.name ? { ...cell, surfaced: true } : cell,
        ]),
      );
    });
    setIsDirty(true);
  }, [cells, isLoaded]);

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
    if (editingAddress && editingAddress !== address && draftEntry.trim().startsWith("=")) {
      const target = getCell(cells, address);
      const reference = target.name || address;
      setDraftEntry((current) => `${current}${current.endsWith("=") || current.endsWith(" ") ? "" : " "}${reference}`);
      return;
    }
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

  function saveConfiguration() {
    const name = configName.trim() || "Untitled Configuration";
    const updatedAt = new Date().toISOString();
    let nextActiveConfigId = activeConfigId;

    setConfigurations((current) => {
      const existing = current.find((configuration) => configuration.id === activeConfigId);
      if (!existing) {
        const created = makeConfiguration(name, cells, columnCount, rowCount);
        nextActiveConfigId = created.id;
        return [...current, created];
      }

      return current.map((configuration) => {
        if (configuration.id !== activeConfigId) return configuration;
        return { ...configuration, name, cells, columnCount, rowCount, updatedAt };
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
    const created = makeConfiguration("Untitled Configuration", {}, defaultColumnCount, defaultRowCount);
    setConfigurations((current) => [...current, created]);
    loadConfiguration(created);
  }

  function duplicateConfiguration() {
    const created = makeConfiguration(`${configName.trim() || "Configuration"} Copy`, cells, columnCount, rowCount);
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
    setActiveConfigId(configuration.id);
    setConfigName(configuration.name);
    setCells(configuration.cells);
    setColumnCount(configuration.columnCount ?? defaultColumnCount);
    setRowCount(configuration.rowCount ?? defaultRowCount);
    setUndoStack([]);
    setRedoStack([]);
    setSelectedAddress("A1");
    setEditingAddress(null);
    setDraftEntry("");
    setIsDirty(false);
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
    setLastImportReport(null);

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
      setImportMessage(`Read ${workbook.sheets.length} worksheet${workbook.sheets.length === 1 ? "" : "s"} from ${file.name}.`);
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

    const converted = convertImportedSheetToQuoin(selectedSheet, { names: pendingImport.names });
    const reviewItems = [...pendingImport.reviewItems, ...converted.reviewItems];
    const importedFormulaCount = selectedSheet.cells.filter((cell) => cell.kind === "formula").length;
    const configurationName = makeImportedConfigurationName(pendingImport.fileName, selectedSheet.name);
    const created = makeConfiguration(configurationName, converted.cells, converted.columnCount, converted.rowCount);

    setConfigurations((current) => [...current, created]);
    loadConfiguration(created);
    setPendingImport(null);
    setImportError("");
    setLastImportReport({
      fileName: pendingImport.fileName,
      sheetName: selectedSheet.name,
      cellCount: Object.keys(converted.cells).length,
      formulaCount: importedFormulaCount,
      promotedNameCount: converted.promotedNameCount,
      reviewItems,
    });
    setImportMessage(
      `Imported ${selectedSheet.name}: ${Object.keys(converted.cells).length} cells, ${importedFormulaCount} formulas, ${converted.promotedNameCount} named cells, ${reviewItems.length} review items.`,
    );
    setActiveView("sheet");
  }

  function cancelPendingImport() {
    setPendingImport(null);
    setImportError("");
  }

  const selectedIssues = issueMap.get(selectedAddress) ?? [];
  const surfacedCells = useMemo(
    () => Object.values(cells).filter((cell) => cell.name && cell.surfaced),
    [cells],
  );
  const selectedImportSheet = pendingImport?.sheets.find((sheet) => sheet.id === selectedImportSheetId) ?? pendingImport?.sheets[0] ?? null;
  const importReviewItems = pendingImport && selectedImportSheet
    ? pendingImport.reviewItems.concat(importReviewItemsForSheet(pendingImport.names, selectedImportSheet.name))
    : [];
  const validationStates = result.ruleStates.filter((rule) => {
    const cell = getCell(cells, rule.address);
    return cell.role === "validation" && cell.surfaced;
  });

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

      {(pendingImport || importMessage || importError || lastImportReport) && (
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
                Worksheet
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
                Import as New Configuration
              </button>
            </div>
          )}

          {lastImportReport && !pendingImport && (
            <div className="importReport">
              <div className="importStats">
                <span>{lastImportReport.cellCount} cells</span>
                <span>{lastImportReport.formulaCount} formulas</span>
                <span>{lastImportReport.promotedNameCount} named cells</span>
                <span>{lastImportReport.reviewItems.length} review items</span>
              </div>
              {lastImportReport.reviewItems.length > 0 ? (
                <div className="importReviewList">
                  {lastImportReport.reviewItems.slice(0, 8).map((item, index) => (
                    <p key={`${item.sheetName ?? lastImportReport.sheetName}-${item.address ?? item.name ?? index}-${index}`} data-severity={item.severity}>
                      <strong>{item.address ?? item.name ?? "Workbook"}</strong>
                      <span>{item.message}</span>
                    </p>
                  ))}
                  {lastImportReport.reviewItems.length > 8 && (
                    <p data-severity="info">
                      <strong>More</strong>
                      <span>{lastImportReport.reviewItems.length - 8} additional review items.</span>
                    </p>
                  )}
                </div>
              ) : (
                <p className="importMessage">No import compatibility review items.</p>
              )}
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
            <span>{selectedAddress}</span>
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

          <div className="authoringLayout">
            <section className="spreadsheetFrame" aria-label="Quoin spreadsheet grid">
              <div
                className="spreadsheetGrid"
                style={{
                  gridTemplateColumns: `44px repeat(${columnCount}, minmax(116px, 1fr))`,
                  minWidth: 44 + columnCount * 116,
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
                  displayValues={displayValues}
                      draftEntry={draftEntry}
                      editInputRef={editInputRef}
                      editingAddress={editingAddress}
                      handleCellClick={handleCellClick}
                      handleGridKeyDown={handleGridKeyDown}
                      handleGridPaste={handleGridPaste}
                      issueMap={issueMap}
                      key={rowNumber}
                      rowNumber={rowNumber}
                      selectedAddress={selectedAddress}
                      setDraftEntry={setDraftEntry}
                      startEditing={startEditing}
                    />
                  );
                })}
              </div>
            </section>

            <Inspector
              clearCell={clearCell}
              dependencySummary={dependencySummary}
              displayValue={displayValues[selectedAddress] ?? ""}
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
          cells={cells}
          displayValues={displayValues}
          result={result}
          surfacedCells={surfacedCells}
          updateCell={updateCell}
          validationStates={validationStates}
        />
      ) : (
        <HelpPanel />
      )}

      <section className="runnerStrip">
        <div>
          <span>Surfaced Outputs</span>
          <strong>{formatOutputs(result.outputs)}</strong>
        </div>
        <div>
          <span>Compliance</span>
          <strong>{result.warnings.length ? result.warnings.map((warning) => warning.message).join(" ") : "No warnings"}</strong>
        </div>
      </section>
    </>
  );
}

function Inspector({
  clearCell,
  dependencySummary,
  displayValue,
  selectedAddress,
  selectedCell,
  selectedIssues,
  updateCell,
  updateLookup,
}: {
  clearCell: (address: string) => void;
  dependencySummary: DependencySummary;
  displayValue: CellValue;
  selectedAddress: string;
  selectedCell: GridCell;
  selectedIssues: string[];
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
                surfaced: name ? selectedCell.surfaced || !selectedCell.name : false,
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
              Type
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

          {selectedCell.role === "input" && (
            <label>
              Dropdown Options
              <textarea
                value={selectedCell.inputOptions.join("\n")}
                onChange={(event) => updateCell(selectedAddress, { inputOptions: splitInputOptions(event.target.value) })}
                placeholder="One option per line. Leave blank for free text."
                rows={3}
              />
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
  cells,
  displayValues,
  result,
  surfacedCells,
  updateCell,
  validationStates,
}: {
  cells: Record<string, GridCell>;
  displayValues: Record<string, CellValue>;
  result: ReturnType<typeof executeEngine>;
  surfacedCells: GridCell[];
  updateCell: (address: string, patch: Partial<GridCell>) => void;
  validationStates: Array<{ address: string; state: string; name?: string | null }>;
}) {
  const inputs = surfacedCells.filter((cell) => cell.role === "input");
  const outputs = surfacedCells.filter((cell) => cell.role !== "input" && cell.role !== "validation" && cell.role !== "compliance" && cell.role !== "action");
  const actions = surfacedCells.filter((cell) => cell.role === "action");
  const surfacedAddresses = new Set(surfacedCells.map((cell) => cell.address));
  const warnings = result.warnings.filter((warning) => surfacedAddresses.has(warning.address));

  return (
    <section className="runnerPreview">
      <div className="runnerHeader">
        <div>
          <p className="eyebrow">Runner Preview</p>
          <h2>Generated Form</h2>
        </div>
        <span data-valid={result.valid}>{result.valid ? "Ready" : "Failed Validation"}</span>
      </div>

      <div className="runnerGrid">
        <div className="runnerPanel">
          <h3>Inputs</h3>
          {inputs.length === 0 ? (
            <p className="runnerEmpty">No surfaced inputs.</p>
          ) : (
            inputs.map((cell) => (
              <label key={cell.address}>
                {labelForCell(cell)}
                {cell.inputOptions.length > 0 ? (
                  <select value={cell.entry} onChange={(event) => updateCell(cell.address, { entry: event.target.value })}>
                    {cell.inputOptions.map((option) => (
                      <option key={option} value={option}>{prettifyName(option)}</option>
                    ))}
                  </select>
                ) : (
                  <input value={cell.entry} onChange={(event) => updateCell(cell.address, { entry: event.target.value })} />
                )}
                {cell.annotation && <small>{cell.annotation}</small>}
              </label>
            ))
          )}
        </div>

        <div className="runnerPanel">
          <h3>Outputs</h3>
          {outputs.length === 0 ? (
            <p className="runnerEmpty">No surfaced outputs.</p>
          ) : (
            outputs.map((cell) => (
              <div className="runnerResult" key={cell.address}>
                <span>{labelForCell(cell)}</span>
                <strong>{formatCellValue(displayValues[cell.address] ?? null)}</strong>
                {cell.annotation && <small>{cell.annotation}</small>}
              </div>
            ))
          )}
        </div>
      </div>

      <div className="runnerMessages">
        <div>
          <h3>Shop Actions</h3>
          {actions.length === 0 ? (
            <p className="runnerEmpty">No shop actions.</p>
          ) : (
            actions.map((cell) => (
              <p data-state="action" key={cell.address}>
                <strong>ACTION</strong>
                {formatCellValue(displayValues[cell.address] ?? null) || labelForCell(cell)}
              </p>
            ))
          )}
        </div>
        <div>
          <h3>Review Flags</h3>
          {warnings.length === 0 ? (
            <p className="runnerEmpty">No review flags.</p>
          ) : (
            warnings.map((warning) => (
              <p data-state="warn" key={warning.cellId}>
                <strong>WARN</strong>
                {warning.message}
              </p>
            ))
          )}
        </div>
      </div>

      <div className="runnerMessages">
        <div>
          <h3>Validation</h3>
          {validationStates.length === 0 ? (
            <p className="runnerEmpty">No validation rules.</p>
          ) : (
            validationStates.map((rule) => {
              const cell = getCell(cells, rule.address);
              return (
                <p data-state={rule.state} key={rule.address}>
                  <strong>{formatCellValue(displayValues[rule.address] ?? null)}</strong>
                  {cell.ruleMessage || cell.annotation || labelForCell(cell)}
                </p>
              );
            })
          )}
        </div>
      </div>
    </section>
  );
}

function HelpPanel() {
  return (
    <section className="helpPanel" aria-label="Quoin help">
      <div className="helpHeader">
        <p className="eyebrow">Help</p>
        <h2>How Quoin Works</h2>
      </div>

      <div className="helpGrid">
        <article>
          <h3>What Quoin Is</h3>
          <p>Quoin starts as a normal spreadsheet, then lets important cells become structured Smart Cells for controlled runner workflows.</p>
        </article>

        <article>
          <h3>Normal Cells</h3>
          <p>Normal cells use coordinates like B2. They can hold values or formulas, but they do not appear in Runner Preview.</p>
        </article>

        <article>
          <h3>Smart Cells</h3>
          <p>Name a cell to promote it. The Smart Cell Name stays formula-safe, while Display Label controls the human-facing runner text.</p>
        </article>

        <article>
          <h3>Runner Preview</h3>
          <p>Surfaced Smart Cells become runner inputs, recommendations, shop actions, validation checks, and review flags.</p>
        </article>

        <article>
          <h3>Roles</h3>
          <p>Use input, formula, output, lookup, action, validation, and compliance roles to describe what each Smart Cell does.</p>
        </article>

        <article>
          <h3>Validation vs Compliance</h3>
          <p>Validation is PASS or FAIL. Compliance is OK or WARN. Warnings do not invalidate the calculation.</p>
        </article>

        <article>
          <h3>Local Configurations</h3>
          <p>New, Save, Duplicate, and Delete manage browser-local configurations. Database-backed publishing comes later.</p>
        </article>

        <article>
          <h3>Beam Demo</h3>
          <p>The demo uses fake lookup data to show the workflow shape: enter drawing conditions, get a recommendation and shop notes.</p>
        </article>
      </div>
    </section>
  );
}

function Row({
  columns,
  cells,
  cellRefs,
  displayValues,
  draftEntry,
  editInputRef,
  editingAddress,
  handleCellClick,
  handleGridKeyDown,
  handleGridPaste,
  issueMap,
  rowNumber,
  selectedAddress,
  setDraftEntry,
  startEditing,
}: {
  columns: string[];
  cells: Record<string, GridCell>;
  cellRefs: React.MutableRefObject<Record<string, HTMLDivElement | null>>;
  displayValues: Record<string, CellValue>;
  draftEntry: string;
  editInputRef: React.RefObject<HTMLInputElement | null>;
  editingAddress: string | null;
  handleCellClick: (address: string) => void;
  handleGridKeyDown: (event: React.KeyboardEvent<HTMLDivElement>, address: string) => void;
  handleGridPaste: (event: React.ClipboardEvent<HTMLDivElement>, address: string) => void;
  issueMap: Map<string, string[]>;
  rowNumber: number;
  selectedAddress: string;
  setDraftEntry: (value: string) => void;
  startEditing: (address: string, replacement?: string) => void;
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
        return (
          <div
            className="gridCell"
            data-editing={editing}
            data-issue={issues.length > 0}
            data-role={cell.name ? cell.role : "normal"}
            data-selected={selected}
            data-smart={Boolean(cell.name)}
            key={address}
            onClick={() => handleCellClick(address)}
            onDoubleClick={() => startEditing(address)}
            onKeyDown={(event) => handleGridKeyDown(event, address)}
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
                onChange={(event) => setDraftEntry(event.target.value)}
              />
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
  return {
    address,
    entry,
    name: "",
    label: "",
    role: entry.startsWith("=") ? "formula" : "input",
    type,
    inputOptions: [],
    surfaced: false,
    annotation: "",
    ruleMessage: "",
    ...options,
  };
}

function hydrateCells(cells: Record<string, GridCell>): Record<string, GridCell> {
  return Object.fromEntries(
    Object.entries(cells).map(([address, cell]) => {
      const hydrated = {
        ...makeCell(address, "", "text"),
        ...cell,
        inputOptions: cell.inputOptions ?? [],
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
    .map((configuration) => ({
      id: configuration.id || makeConfigId(),
      name: configuration.name || "Untitled Configuration",
      cells: hydrateCells(configuration.cells ?? {}),
      columnCount: configuration.columnCount ?? defaultColumnCount,
      rowCount: configuration.rowCount ?? defaultRowCount,
      updatedAt: configuration.updatedAt || new Date().toISOString(),
    }));
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

function makeConfiguration(name: string, cells: Record<string, GridCell>, columnCount = defaultColumnCount, rowCount = defaultRowCount): LocalConfiguration {
  return {
    id: makeConfigId(),
    name,
    cells: hydrateCells(cells),
    columnCount,
    rowCount,
    updatedAt: new Date().toISOString(),
  };
}

function makeImportedConfigurationName(fileName: string, sheetName: string): string {
  const baseName = fileName.replace(/\.[^.]+$/, "").trim() || "Workbook";
  return `Imported - ${baseName} - ${sheetName}`;
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

function getCell(cells: Record<string, GridCell>, address: string): GridCell {
  return cells[address] ?? makeCell(address, "", "text");
}

function applyCellPatch(cell: GridCell, patch: Partial<GridCell>): GridCell {
  const next = { ...cell, ...patch };
  const hasName = Boolean(next.name);
  const hasFormulaEntry = typeof next.entry === "string" && next.entry.trim().startsWith("=");

  if (hasName && hasFormulaEntry && next.role === "input" && patch.role === undefined) {
    next.role = "formula";
  }

  return next;
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
