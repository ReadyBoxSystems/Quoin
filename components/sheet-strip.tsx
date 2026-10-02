"use client";

import type { WorkbookSheet } from "@/lib/sheet/types";

export function SheetStrip({ activeSheetId, addSheet, renameSheet, sheets, switchSheet }: {
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
          <button key={sheet.id} type="button" data-active={sheet.id === activeSheetId}
            onClick={() => switchSheet(sheet.id)} role="tab" aria-selected={sheet.id === activeSheetId} title={sheet.name}>
            <span>{sheet.name}</span>
            <small>{Object.keys(sheet.cells).length} cells</small>
          </button>
        ))}
      </div>
      <button type="button" className="sheetAddButton" onClick={addSheet}>Add Sheet</button>
      {activeSheet && (
        <label className="sheetRename">
          <span>Active Sheet</span>
          <input value={activeSheet.name} onChange={(event) => renameSheet(activeSheet.id, event.target.value)} />
        </label>
      )}
    </aside>
  );
}
