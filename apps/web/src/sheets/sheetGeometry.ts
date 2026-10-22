// Builds the sheet's geometry (row heights, column widths, hidden lines,
// merged cells, frozen panes) from the engine, adding automatic row heights
// the way Excel autofits rows: a row grows to fit wrapped text, line breaks
// and big fonts unless the user gave it a height.

import { parseStyle, type SheetsDocument } from "@office/engine";
import { cellLines, lineHeightFor } from "./cellRenderer";
import { ColumnLayout, DEFAULT_ROW_HEIGHT, RowLayout, type Run } from "./layout";
import { MergeIndex } from "./merges";
import type { SheetGeometry } from "./paint";
import { NO_FREEZE } from "./panes";

/** Excel's tallest row: 409 points. */
export const MAX_ROW_HEIGHT = Math.floor((409 * 4) / 3);

let measureCtx: CanvasRenderingContext2D | null = null;
function measurer(): CanvasRenderingContext2D {
  measureCtx ??= document.createElement("canvas").getContext("2d")!;
  return measureCtx;
}

/** Height row `row` needs to show all its content, or the default height. */
export function autoRowHeight(doc: SheetsDocument, sheet: number, row: number, cols: ColumnLayout, merges: MergeIndex): number {
  const [, extentCols] = doc.extent(sheet);
  if (extentCols === 0) return DEFAULT_ROW_HEIGHT;
  const texts = doc.displayBlock(sheet, row, 0, 1, extentCols);
  const styles = doc.styleBlock(sheet, row, 0, 1, extentCols);
  const ctx = measurer();
  let needed = DEFAULT_ROW_HEIGHT;
  for (let c = 0; c < extentCols; c++) {
    if (!styles[c] && !texts[c].includes("\n")) continue;
    // Merged cells: Excel grows the row for a bigger font, but never to fit
    // wrapped text, and only for areas that are one row tall.
    const merged = merges.at(row, c);
    if (merged && (merged.top !== merged.bottom || merged.left !== c)) continue;
    const style = parseStyle(styles[c]);
    const text = texts[c];
    // One line of 11pt Calibri is Excel's 20px row; bigger fonts and wrapped lines scale from there.
    const lines = style.wrap && text && !merged ? cellLines(ctx, text, style, cols.width(c)).length : 1;
    const h = lines * lineHeightFor(style);
    needed = Math.max(needed, Math.ceil(h));
  }
  return Math.min(needed, MAX_ROW_HEIGHT);
}

/** The engine's row layout plus automatic heights for rows the user hasn't sized. */
export function buildRows(doc: SheetsDocument, sheet: number, cols: ColumnLayout, merges: MergeIndex): RowLayout {
  const flat = doc.rowLayout(sheet);
  const n = flat[0] ?? 0;
  const custom: Run[] = [];
  for (let k = 0; k < n; k++) custom.push([flat[1 + 3 * k], flat[2 + 3 * k], flat[3 + 3 * k]]);
  const hidden: Array<[number, number]> = [];
  for (let i = 1 + 3 * n; i + 1 < flat.length; i += 2) hidden.push([flat[i], flat[i + 1]]);

  const isCustom = (row: number) => custom.some(([a, b]) => row >= a && row <= b);
  const auto: Run[] = [];
  for (const row of doc.rowsToMeasure(sheet)) {
    if (isCustom(row)) continue;
    const h = autoRowHeight(doc, sheet, row, cols, merges);
    if (h !== DEFAULT_ROW_HEIGHT) auto.push([row, row, h]);
  }
  // Later runs win, so a height the user set beats the automatic one.
  return new RowLayout([...auto, ...custom], hidden);
}

export function buildGeometry(doc: SheetsDocument, sheet: number): SheetGeometry {
  const cols = new ColumnLayout(doc.columnWidths(sheet), doc.hiddenColumns(sheet));
  const merges = new MergeIndex(doc.merges(sheet));
  const rows = buildRows(doc, sheet, cols, merges);
  const f = doc.freezePanes(sheet);
  return { rows, cols, merges, freeze: f.length === 4 ? { rows: f[0], cols: f[1], top: f[2], left: f[3] } : NO_FREEZE };
}
