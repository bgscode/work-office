// Paints the visible part of the grid onto a canvas. Pure function of its
// inputs: Grid.tsx calls it whenever anything visible changes.

import { parseStyle, type SheetsDocument } from "@office/engine";
import { drawBorders, drawCell, spillClip, textOverflow } from "./cellRenderer";
import { HEADER_HEIGHT, HEADER_WIDTH, MAX_COLS, MAX_ROWS, ROW_HEIGHT, colName, selectionRect, type Rect, type Selection } from "./geometry";
import type { ColumnLayout } from "./layout";

const theme = {
  background: "#ffffff",
  gridLine: "#e2e2e2",
  text: "#1f1f1f",
  error: "#c5221f",
  headerBg: "#f4f5f7",
  headerActiveBg: "#d3e3fd",
  headerText: "#555b66",
  selectionFill: "rgba(24, 90, 189, 0.10)",
  selectionBorder: "#185abd",
  fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif",
};

/** Colours for references while a formula is being edited (cycled in order). */
export const REF_COLORS = ["#1a73e8", "#e8710a", "#9334e6", "#188038", "#d93025", "#12b5cb"];

/** Size of the fill handle square at the bottom-right of the selection. */
export const FILL_HANDLE_SIZE = 7;

export interface PaintInput {
  width: number;
  height: number;
  scroll: { top: number; left: number };
  doc: SheetsDocument;
  sheet: number;
  selection: Selection;
  layout: ColumnLayout;
  /** References in the formula being edited, drawn as coloured boxes. */
  highlights: Rect[];
  /** Area a fill-handle drag would fill. */
  fillPreview: Rect | null;
  showFillHandle: boolean;
}

/** Viewport x of a column's left edge. */
export function colX(layout: ColumnLayout, col: number, scrollLeft: number): number {
  return HEADER_WIDTH + layout.x(col) - scrollLeft;
}

/** Viewport y of a row's top edge. */
export function rowY(row: number, scrollTop: number): number {
  return HEADER_HEIGHT + row * ROW_HEIGHT - scrollTop;
}

export function paint(ctx: CanvasRenderingContext2D, p: PaintInput) {
  const { width, height, scroll, doc, sheet, selection, layout } = p;
  const firstRow = Math.floor(scroll.top / ROW_HEIGHT);
  const rows = Math.min(Math.ceil((height - HEADER_HEIGHT) / ROW_HEIGHT) + 1, MAX_ROWS - firstRow);
  const firstCol = layout.colAt(scroll.left);
  let cols = 0;
  while (firstCol + cols < MAX_COLS && colX(layout, firstCol + cols, scroll.left) < width) cols++;

  const xOf = (c: number) => colX(layout, c, scroll.left);
  const yOf = (r: number) => rowY(r, scroll.top);
  const rectPx = (r: Rect) => ({
    x: xOf(r.left),
    y: yOf(r.top),
    w: xOf(r.right + 1) - xOf(r.left),
    h: (r.bottom - r.top + 1) * ROW_HEIGHT,
  });
  const sel = selectionRect(selection);

  const texts = doc.displayBlock(sheet, firstRow, firstCol, rows, cols);
  const kinds = doc.kindBlock(sheet, firstRow, firstCol, rows, cols);
  const styles = doc.styleBlock(sheet, firstRow, firstCol, rows, cols).map(parseStyle);
  const at = (r: number, c: number) => r * cols + c;

  ctx.fillStyle = theme.background;
  ctx.fillRect(0, 0, width, height);

  // Everything cell-related is clipped so it never draws over the headers.
  ctx.save();
  ctx.beginPath();
  ctx.rect(HEADER_WIDTH, HEADER_HEIGHT, width - HEADER_WIDTH, height - HEADER_HEIGHT);
  ctx.clip();

  // Cell background colours.
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const fill = styles[at(r, c)].fill;
      if (!fill) continue;
      ctx.fillStyle = fill;
      ctx.fillRect(xOf(firstCol + c), yOf(firstRow + r), layout.width(firstCol + c), ROW_HEIGHT);
    }
  }

  // Grid lines.
  ctx.strokeStyle = theme.gridLine;
  ctx.lineWidth = 1;
  ctx.beginPath();
  for (let r = 0; r <= rows; r++) {
    const y = Math.floor(yOf(firstRow + r)) + 0.5;
    ctx.moveTo(HEADER_WIDTH, y);
    ctx.lineTo(width, y);
  }
  for (let c = 0; c <= cols; c++) {
    const x = Math.floor(xOf(firstCol + c)) + 0.5;
    ctx.moveTo(x, HEADER_HEIGHT);
    ctx.lineTo(x, height);
  }
  ctx.stroke();

  // Cell text (cellRenderer.ts), then borders on top so overflowing text never hides them.
  const textOpts = { textColor: theme.text, errorColor: theme.error };
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const i = at(r, c);
      if (!texts[i]) continue;
      const col = firstCol + c;
      const cell = { x: xOf(col), y: yOf(firstRow + r), w: layout.width(col), h: ROW_HEIGHT };
      const overflow = textOverflow(ctx, cell, texts[i], kinds[i], styles[i]);
      const free = (cc: number) => cc >= firstCol && cc < firstCol + cols && !texts[at(r, cc - firstCol)] && !styles[at(r, cc - firstCol)].fill;
      const clip = spillClip(cell, overflow, col, free, (cc) => layout.width(cc));
      drawCell(ctx, cell, texts[i], kinds[i], styles[i], { ...textOpts, clipLeft: clip.left, clipRight: clip.right });
    }
  }
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const style = styles[at(r, c)];
      if (!style.border) continue;
      const col = firstCol + c;
      drawBorders(ctx, { x: xOf(col), y: yOf(firstRow + r), w: layout.width(col), h: ROW_HEIGHT }, style.border, style.fill ?? theme.background);
    }
  }

  // Selection tint, leaving the active cell clear.
  const selPx = rectPx(sel);
  const active = rectPx({ top: selection.anchor.row, left: selection.anchor.col, bottom: selection.anchor.row, right: selection.anchor.col });
  ctx.fillStyle = theme.selectionFill;
  ctx.beginPath();
  ctx.rect(selPx.x, selPx.y, selPx.w, selPx.h);
  ctx.rect(active.x, active.y, active.w, active.h);
  ctx.fill("evenodd");

  // References in the formula being edited.
  p.highlights.forEach((h, n) => {
    const color = REF_COLORS[n % REF_COLORS.length];
    const px = rectPx(h);
    ctx.fillStyle = color + "1a";
    ctx.fillRect(px.x, px.y, px.w, px.h);
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    ctx.strokeRect(px.x + 1, px.y + 1, px.w - 1, px.h - 1);
  });

  // Selection border and fill handle.
  ctx.strokeStyle = theme.selectionBorder;
  ctx.lineWidth = 2;
  ctx.strokeRect(selPx.x + 1, selPx.y + 1, selPx.w - 1, selPx.h - 1);
  if (p.showFillHandle) {
    const hx = selPx.x + selPx.w;
    const hy = selPx.y + selPx.h;
    const s = FILL_HANDLE_SIZE;
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(hx - s / 2 - 1, hy - s / 2 - 1, s + 2, s + 2);
    ctx.fillStyle = theme.selectionBorder;
    ctx.fillRect(hx - s / 2, hy - s / 2, s, s);
  }
  if (p.fillPreview) {
    const px = rectPx(p.fillPreview);
    ctx.setLineDash([4, 3]);
    ctx.strokeStyle = "#5f6368";
    ctx.lineWidth = 1;
    ctx.strokeRect(px.x + 0.5, px.y + 0.5, px.w - 1, px.h - 1);
    ctx.setLineDash([]);
  }
  ctx.restore();

  paintHeaders(ctx, p, firstRow, rows, firstCol, cols, sel);
}

function paintHeaders(
  ctx: CanvasRenderingContext2D,
  p: PaintInput,
  firstRow: number,
  rows: number,
  firstCol: number,
  cols: number,
  sel: Rect,
) {
  const { width, height, scroll, layout } = p;
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(0, 0, width, HEADER_HEIGHT);
  ctx.fillRect(0, 0, HEADER_WIDTH, height);
  ctx.font = `13px ${theme.fontFamily}`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  ctx.save();
  ctx.beginPath();
  ctx.rect(HEADER_WIDTH, 0, width - HEADER_WIDTH, HEADER_HEIGHT);
  ctx.clip();
  ctx.strokeStyle = theme.gridLine;
  ctx.lineWidth = 1;
  for (let c = 0; c < cols; c++) {
    const col = firstCol + c;
    const x = colX(layout, col, scroll.left);
    const w = layout.width(col);
    if (col >= sel.left && col <= sel.right) {
      ctx.fillStyle = theme.headerActiveBg;
      ctx.fillRect(x, 0, w, HEADER_HEIGHT);
    }
    ctx.fillStyle = theme.headerText;
    ctx.fillText(colName(col), x + w / 2, HEADER_HEIGHT / 2);
    ctx.beginPath();
    ctx.moveTo(Math.floor(x + w) + 0.5, 0);
    ctx.lineTo(Math.floor(x + w) + 0.5, HEADER_HEIGHT);
    ctx.stroke();
  }
  ctx.restore();

  ctx.save();
  ctx.beginPath();
  ctx.rect(0, HEADER_HEIGHT, HEADER_WIDTH, height - HEADER_HEIGHT);
  ctx.clip();
  for (let r = 0; r < rows; r++) {
    const row = firstRow + r;
    const y = rowY(row, scroll.top);
    if (row >= sel.top && row <= sel.bottom) {
      ctx.fillStyle = theme.headerActiveBg;
      ctx.fillRect(0, y, HEADER_WIDTH, ROW_HEIGHT);
    }
    ctx.fillStyle = theme.headerText;
    ctx.fillText(String(row + 1), HEADER_WIDTH / 2, y + ROW_HEIGHT / 2);
  }
  ctx.restore();
  ctx.textAlign = "start";

  ctx.strokeStyle = theme.gridLine;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(0, HEADER_HEIGHT - 0.5);
  ctx.lineTo(width, HEADER_HEIGHT - 0.5);
  ctx.moveTo(HEADER_WIDTH - 0.5, 0);
  ctx.lineTo(HEADER_WIDTH - 0.5, height);
  ctx.stroke();
}

export { fontFor } from "./cellRenderer";
