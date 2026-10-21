// Paints the visible part of the grid onto a canvas. Pure function of its
// inputs: Grid.tsx calls it whenever anything visible changes.

import { ValueKind, parseStyle, type CellStyle, type SheetsDocument } from "@office/engine";
import { HEADER_HEIGHT, HEADER_WIDTH, MAX_COLS, MAX_ROWS, ROW_HEIGHT, colName, selectionRect, type Rect, type Selection } from "./geometry";
import type { ColumnLayout } from "./layout";

// Excel (Microsoft 365) colours and fonts.
const theme = {
  background: "#ffffff",
  gridLine: "#e1e1e1",
  text: "#000000",
  error: "#c00000",
  headerBg: "#f5f5f5",
  headerBorder: "#d4d4d4",
  headerActiveBg: "#e1e1e1",
  /** Green line on the selected rows'/columns' header edge. */
  headerActiveLine: "#217346",
  headerActiveText: "#217346",
  headerText: "#444444",
  selectionFill: "rgba(0, 0, 0, 0.08)",
  selectionBorder: "#217346",
  fontFamily: "Calibri, Carlito, 'Segoe UI', Arial, sans-serif",
};

/** Default cell font size in px (Calibri 11pt at 96 dpi). */
export const CELL_FONT_SIZE = 15;

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

  // Cell text.
  ctx.textBaseline = "middle";
  const pad = 4;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const i = at(r, c);
      // Line breaks (Alt+Enter) show as spaces until wrap text is supported.
      const text = texts[i].includes("\n") ? texts[i].replace(/\n/g, " ") : texts[i];
      if (!text) continue;
      const kind = kinds[i];
      const style = styles[i];
      const col = firstCol + c;
      const x = xOf(col);
      const y = yOf(firstRow + r);
      const w = layout.width(col);
      ctx.font = fontFor(style);
      const align = style.align ?? (kind === ValueKind.Number ? "right" : kind === ValueKind.Text ? "left" : "center");
      let shown = text;
      let textWidth = ctx.measureText(shown).width;

      // Numbers that don't fit show ### like Excel, rather than a misleading partial number.
      if (kind === ValueKind.Number && textWidth > w - 2 * pad) {
        shown = "#".repeat(Math.max(1, Math.floor((w - 2 * pad) / ctx.measureText("#").width)));
        textWidth = ctx.measureText(shown).width;
      }

      // Left-aligned text may spill into empty cells to the right.
      let clipRight = x + w;
      if (align === "left" && textWidth > w - 2 * pad) {
        let next = c + 1;
        while (next < cols && !texts[at(r, next)] && !styles[at(r, next)].fill && clipRight < x + pad + textWidth) {
          clipRight += layout.width(firstCol + next);
          next++;
        }
        if (clipRight > x + w) {
          ctx.fillStyle = theme.background;
          ctx.fillRect(x + w, y + 1, clipRight - (x + w), ROW_HEIGHT - 1);
        }
      }

      const tx = align === "right" ? x + w - pad - textWidth : align === "center" ? x + (w - textWidth) / 2 : x + pad;
      ctx.fillStyle = kind === ValueKind.Error ? theme.error : style.color ?? theme.text;
      ctx.save();
      ctx.beginPath();
      ctx.rect(x + 1, y, clipRight - x - 2, ROW_HEIGHT);
      ctx.clip();
      ctx.fillText(shown, tx, y + ROW_HEIGHT / 2);
      if (style.u || style.s) {
        ctx.strokeStyle = ctx.fillStyle;
        ctx.lineWidth = 1;
        ctx.beginPath();
        if (style.u) {
          ctx.moveTo(tx, y + ROW_HEIGHT / 2 + 7.5);
          ctx.lineTo(tx + textWidth, y + ROW_HEIGHT / 2 + 7.5);
        }
        if (style.s) {
          ctx.moveTo(tx, y + ROW_HEIGHT / 2 + 0.5);
          ctx.lineTo(tx + textWidth, y + ROW_HEIGHT / 2 + 0.5);
        }
        ctx.stroke();
      }
      ctx.restore();
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
  ctx.strokeStyle = theme.headerBorder;
  ctx.lineWidth = 1;
  const headerText = (inSel: boolean) => (inSel ? theme.headerActiveText : theme.headerText);
  for (let c = 0; c < cols; c++) {
    const col = firstCol + c;
    const x = colX(layout, col, scroll.left);
    const w = layout.width(col);
    const inSel = col >= sel.left && col <= sel.right;
    if (inSel) {
      ctx.fillStyle = theme.headerActiveBg;
      ctx.fillRect(x, 0, w, HEADER_HEIGHT);
      ctx.fillStyle = theme.headerActiveLine;
      ctx.fillRect(x, HEADER_HEIGHT - 2, w, 2);
    }
    ctx.fillStyle = headerText(inSel);
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
    const inSel = row >= sel.top && row <= sel.bottom;
    if (inSel) {
      ctx.fillStyle = theme.headerActiveBg;
      ctx.fillRect(0, y, HEADER_WIDTH, ROW_HEIGHT);
      ctx.fillStyle = theme.headerActiveLine;
      ctx.fillRect(HEADER_WIDTH - 2, y, 2, ROW_HEIGHT);
    }
    ctx.fillStyle = headerText(inSel);
    ctx.fillText(String(row + 1), HEADER_WIDTH / 2, y + ROW_HEIGHT / 2);
    // Excel separates row numbers with lines, like the column headers.
    ctx.fillStyle = theme.headerBorder;
    ctx.fillRect(0, Math.floor(y + ROW_HEIGHT), HEADER_WIDTH, 1);
  }
  ctx.restore();
  ctx.textAlign = "start";

  ctx.strokeStyle = theme.headerBorder;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(0, HEADER_HEIGHT - 0.5);
  ctx.lineTo(width, HEADER_HEIGHT - 0.5);
  ctx.moveTo(HEADER_WIDTH - 0.5, 0);
  ctx.lineTo(HEADER_WIDTH - 0.5, height);
  ctx.stroke();
}

export function fontFor(style: CellStyle): string {
  return `${style.i ? "italic " : ""}${style.b ? "bold " : ""}${CELL_FONT_SIZE}px ${theme.fontFamily}`;
}
