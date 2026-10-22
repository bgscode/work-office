// Paints the visible part of the grid onto a canvas. Pure function of its
// inputs: Grid.tsx calls it whenever anything visible changes.
//
// The viewport is split into panes (panes.ts): with frozen rows/columns there
// are up to four, each clipped and scrolled on its own. Row heights, column
// widths and hidden lines come from layout.ts; merged cells from merges.ts;
// what goes inside one cell (text, wrap, borders) from cellRenderer.ts.

import { parseStyle, type CellStyle, type SheetsDocument } from "@office/engine";
import { drawBorders, drawCell, spillClip, textOverflow } from "./cellRenderer";
import { HEADER_HEIGHT, HEADER_WIDTH, MAX_COLS, MAX_ROWS, colName, selectionRect, type Rect, type Selection } from "./geometry";
import type { FilterInfo } from "./features/dataTools";
import type { ColumnLayout, RowLayout } from "./layout";
import type { MergeIndex } from "./merges";
import { Panes, type Freeze, type Pane } from "./panes";

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
  /** Marks hidden rows/columns in the headers. */
  hiddenMarker: "#8a8886",
  freezeLine: "#9e9e9e",
  /** Row numbers of rows a filter is narrowing, and filter buttons that filter. */
  filtered: "#0063b1",
  /** Circle Invalid Data. */
  invalidCircle: "#e00000",
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

/** Everything about the sheet's shape that painting and hit-testing need. */
export interface SheetGeometry {
  rows: RowLayout;
  cols: ColumnLayout;
  merges: MergeIndex;
  freeze: Freeze;
}

export interface PaintInput {
  width: number;
  height: number;
  scroll: { top: number; left: number };
  doc: SheetsDocument;
  sheet: number;
  selection: Selection;
  geo: SheetGeometry;
  /** References in the formula being edited, drawn as coloured boxes. */
  highlights: Rect[];
  /** Area a fill-handle drag would fill. */
  fillPreview: Rect | null;
  showFillHandle: boolean;
  /** AutoFilter: dropdown buttons on its header row. */
  filter: FilterInfo | null;
  /** Circle Invalid Data. */
  circles: { row: number; col: number }[];
  /** The dynamic array the active cell belongs to: outlined in blue, like Excel. */
  spill: Rect | null;
}

/** Size of an AutoFilter dropdown button. */
export const FILTER_BUTTON_SIZE = 16;

/** Where the filter button sits in a header cell: bottom-right, like Excel. */
export function filterButtonBox(cell: { x: number; y: number; w: number; h: number }) {
  const size = Math.min(FILTER_BUTTON_SIZE, cell.h - 2, cell.w - 2);
  return { x: cell.x + cell.w - size - 2, y: cell.y + cell.h - size - 2, size };
}

export function makePanes(geo: SheetGeometry, scroll: { top: number; left: number }, view: { width: number; height: number }): Panes {
  return new Panes(geo.rows, geo.cols, geo.freeze, scroll, view);
}

/** Coordinates of lines inside one pane. */
function paneMap(geo: SheetGeometry, pane: Pane) {
  const x = (c: number) => HEADER_WIDTH + geo.cols.x(Math.min(c, MAX_COLS)) - pane.cols.offset;
  const y = (r: number) => HEADER_HEIGHT + geo.rows.y(Math.min(r, MAX_ROWS)) - pane.rows.offset;
  const rect = (r: Rect) => ({ x: x(r.left), y: y(r.top), w: x(r.right + 1) - x(r.left), h: y(r.bottom + 1) - y(r.top) });
  return { x, y, rect };
}

export function paint(ctx: CanvasRenderingContext2D, p: PaintInput) {
  const { width, height, geo } = p;
  const panes = makePanes(geo, p.scroll, { width, height });

  ctx.fillStyle = theme.background;
  ctx.fillRect(0, 0, width, height);

  for (const pane of panes.panes()) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(pane.cols.from, pane.rows.from, pane.cols.to - pane.cols.from, pane.rows.to - pane.rows.from);
    ctx.clip();
    paintCells(ctx, p, pane);
    paintOverlays(ctx, p, pane);
    ctx.restore();
  }

  // Excel's thin line under frozen rows / right of frozen columns.
  ctx.fillStyle = theme.freezeLine;
  if (geo.freeze.rows > 0) ctx.fillRect(HEADER_WIDTH, HEADER_HEIGHT + panes.frozenHeight - 1, width - HEADER_WIDTH, 1);
  if (geo.freeze.cols > 0) ctx.fillRect(HEADER_WIDTH + panes.frozenWidth - 1, HEADER_HEIGHT, 1, height - HEADER_HEIGHT);

  paintHeaders(ctx, p, panes);
}

function paintCells(ctx: CanvasRenderingContext2D, p: PaintInput, pane: Pane) {
  const { doc, sheet, geo } = p;
  const { rows, cols, merges } = geo;
  const { x, y, rect } = paneMap(geo, pane);
  const r0 = pane.rows.first;
  const c0 = pane.cols.first;
  const nRows = pane.rows.last - r0 + 1;
  const nCols = pane.cols.last - c0 + 1;
  if (nRows <= 0 || nCols <= 0) return;

  const texts = doc.displayBlock(sheet, r0, c0, nRows, nCols);
  const kinds = doc.kindBlock(sheet, r0, c0, nRows, nCols);
  const styles = doc.styleBlock(sheet, r0, c0, nRows, nCols).map(parseStyle);
  // Conditional formatting overrides the cell's own fill, colour and font style;
  // data bars are drawn under the text.
  const cf = doc.cfBlock(sheet, r0, c0, nRows, nCols);
  const bars: ({ len: number; color: string } | null)[] = [];
  for (let i = 0; i < cf.length; i++) {
    if (!cf[i]) continue;
    const { bar, barColor, ...rest } = JSON.parse(cf[i]) as CellStyle & { bar?: number; barColor?: string };
    styles[i] = { ...styles[i], ...rest };
    if (bar !== undefined && barColor) bars[i] = { len: bar, color: barColor };
  }
  const at = (r: number, c: number) => (r - r0) * nCols + (c - c0);
  const visRows = rows.visibleLines(r0, pane.rows.last);
  const visCols = cols.visibleLines(c0, pane.cols.last);

  // Cells covered by a merged area are drawn once, as the whole area.
  const block: Rect = { top: r0, left: c0, bottom: pane.rows.last, right: pane.cols.last };
  const mergeList = merges.within(block);
  const covered = new Uint8Array(nRows * nCols);
  for (const m of mergeList) {
    for (let r = Math.max(m.top, r0); r <= Math.min(m.bottom, block.bottom); r++) {
      for (let c = Math.max(m.left, c0); c <= Math.min(m.right, block.right); c++) covered[at(r, c)] = 1;
    }
  }

  // Background fills.
  for (const r of visRows) {
    for (const c of visCols) {
      const i = at(r, c);
      if (covered[i] || !styles[i].fill) continue;
      ctx.fillStyle = styles[i].fill!;
      ctx.fillRect(x(c), y(r), cols.width(c), rows.height(r));
    }
  }

  // Data bars: a gradient from the bar colour, as in Excel's default gradient fill.
  if (bars.length) {
    for (const r of visRows) {
      for (const c of visCols) {
        const bar = bars[at(r, c)];
        if (!bar || covered[at(r, c)]) continue;
        const bx = x(c) + 2;
        const by = y(r) + 2;
        const bw = Math.max(0, (cols.width(c) - 5) * bar.len);
        const bh = rows.height(r) - 5;
        if (bw <= 0 || bh <= 0) continue;
        const grad = ctx.createLinearGradient(bx, 0, bx + bw, 0);
        grad.addColorStop(0, bar.color);
        grad.addColorStop(1, bar.color + "33");
        ctx.fillStyle = grad;
        ctx.fillRect(bx, by, bw, bh);
        ctx.strokeStyle = bar.color;
        ctx.lineWidth = 1;
        ctx.strokeRect(bx + 0.5, by + 0.5, bw - 1, bh - 1);
      }
    }
  }

  // Grid lines along every visible row and column edge.
  ctx.strokeStyle = theme.gridLine;
  ctx.lineWidth = 1;
  ctx.beginPath();
  for (const r of visRows) {
    const yy = Math.floor(y(r) + rows.height(r)) - 0.5;
    ctx.moveTo(pane.cols.from, yy);
    ctx.lineTo(pane.cols.to, yy);
  }
  for (const c of visCols) {
    const xx = Math.floor(x(c) + cols.width(c)) - 0.5;
    ctx.moveTo(xx, pane.rows.from);
    ctx.lineTo(xx, pane.rows.to);
  }
  ctx.stroke();

  // Merged areas: one background over the whole area hides the inner grid lines.
  const anchorOf = (m: Rect) => {
    if (m.top >= r0 && m.left >= c0 && m.top <= block.bottom && m.left <= block.right) {
      const i = at(m.top, m.left);
      return { text: texts[i], kind: kinds[i], style: styles[i] };
    }
    return {
      text: doc.displayBlock(sheet, m.top, m.left, 1, 1)[0],
      kind: doc.kindBlock(sheet, m.top, m.left, 1, 1)[0],
      style: parseStyle(doc.styleBlock(sheet, m.top, m.left, 1, 1)[0]),
    };
  };
  const merged = mergeList.map((m) => ({ area: m, px: rect(m), ...anchorOf(m) }));
  for (const m of merged) {
    if (m.px.w <= 0 || m.px.h <= 0) continue;
    ctx.fillStyle = m.style.fill ?? theme.background;
    ctx.fillRect(m.px.x, m.px.y, m.px.w - 1, m.px.h - 1);
  }

  // Cell text, then borders on top so overflowing text never hides them.
  const textOpts = { textColor: theme.text, errorColor: theme.error };
  for (const r of visRows) {
    for (const c of visCols) {
      const i = at(r, c);
      if (covered[i] || !texts[i]) continue;
      const cell = { x: x(c), y: y(r), w: cols.width(c), h: rows.height(r) };
      const overflow = textOverflow(ctx, cell, texts[i], kinds[i], styles[i]);
      const free = (cc: number) => {
        if (cc < c0 || cc > block.right) return false;
        const j = at(r, cc);
        return !covered[j] && !texts[j] && !styles[j].fill;
      };
      const clip = spillClip(cell, overflow, c, free, (cc) => cols.width(cc));
      drawCell(ctx, cell, texts[i], kinds[i], styles[i], { ...textOpts, clipLeft: clip.left, clipRight: clip.right });
    }
  }
  for (const m of merged) {
    if (!m.text || m.px.w <= 0 || m.px.h <= 0) continue;
    drawCell(ctx, m.px, m.text, m.kind, m.style, { ...textOpts, clipLeft: m.px.x, clipRight: m.px.x + m.px.w });
  }
  for (const r of visRows) {
    for (const c of visCols) {
      const i = at(r, c);
      const style: CellStyle = styles[i];
      if (covered[i] || !style.border) continue;
      drawBorders(ctx, { x: x(c), y: y(r), w: cols.width(c), h: rows.height(r) }, style.border, style.fill ?? theme.background);
    }
  }
  for (const m of merged) {
    if (m.style.border) drawBorders(ctx, m.px, m.style.border, m.style.fill ?? theme.background);
  }

  // AutoFilter dropdown buttons on the header row.
  const f = p.filter;
  if (f && f.range.top >= r0 && f.range.top <= block.bottom && !rows.isHidden(f.range.top)) {
    const r = f.range.top;
    for (const c of visCols) {
      if (c < f.range.left || c > f.range.right) continue;
      // A merged header gets one button, at its right edge.
      const area = merges.cellRect({ row: r, col: c });
      const merged = area.left !== area.right || area.top !== area.bottom;
      if (merged && c !== Math.min(area.right, f.range.right)) continue;
      const cell = merged ? rect(area) : { x: x(c), y: y(r), w: cols.width(c), h: rows.height(r) };
      drawFilterButton(ctx, filterButtonBox(cell), f.columns.has(c), f.sort?.col === c ? f.sort.ascending : null);
    }
  }
}

/** Excel's filter button: a small raised box with a dropdown arrow, a funnel when filtering, an arrow when sorted. */
function drawFilterButton(ctx: CanvasRenderingContext2D, b: { x: number; y: number; size: number }, filtering: boolean, sorted: boolean | null) {
  const { x, y, size } = b;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(x, y, size, size);
  ctx.strokeStyle = "#a6a6a6";
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, size - 1, size - 1);
  ctx.fillStyle = filtering ? theme.filtered : "#444444";
  const cx = x + size / 2;
  const cy = y + size / 2;
  if (filtering) {
    // Funnel on the left, small arrow on the right.
    ctx.beginPath();
    ctx.moveTo(x + 3, y + 4);
    ctx.lineTo(x + 10, y + 4);
    ctx.lineTo(x + 7.5, y + 8);
    ctx.lineTo(x + 7.5, y + 12);
    ctx.lineTo(x + 5.5, y + 11);
    ctx.lineTo(x + 5.5, y + 8);
    ctx.closePath();
    ctx.fill();
    ctx.beginPath();
    ctx.moveTo(x + size - 6, cy);
    ctx.lineTo(x + size - 2.5, cy);
    ctx.lineTo(x + size - 4.25, cy + 2.5);
    ctx.closePath();
    ctx.fill();
  } else if (sorted !== null) {
    // Sort arrow next to the dropdown arrow.
    ctx.fillRect(x + 4, y + 4, 1, size - 8);
    ctx.beginPath();
    if (sorted) {
      ctx.moveTo(x + 2.5, y + size - 6);
      ctx.lineTo(x + 6.5, y + size - 6);
      ctx.lineTo(x + 4.5, y + size - 3.5);
    } else {
      ctx.moveTo(x + 2.5, y + 6);
      ctx.lineTo(x + 6.5, y + 6);
      ctx.lineTo(x + 4.5, y + 3.5);
    }
    ctx.closePath();
    ctx.fill();
    ctx.beginPath();
    ctx.moveTo(x + size - 8, cy - 1);
    ctx.lineTo(x + size - 3, cy - 1);
    ctx.lineTo(x + size - 5.5, cy + 2);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.moveTo(cx - 3.5, cy - 1.5);
    ctx.lineTo(cx + 3.5, cy - 1.5);
    ctx.lineTo(cx, cy + 2.5);
    ctx.closePath();
    ctx.fill();
  }
}

function paintOverlays(ctx: CanvasRenderingContext2D, p: PaintInput, pane: Pane) {
  const { selection, geo } = p;
  const { rect } = paneMap(geo, pane);
  const sel = selectionRect(selection);

  // Selection tint, leaving the active cell (or its merged area) clear.
  const selPx = rect(sel);
  const active = rect(geo.merges.cellRect(selection.anchor));
  ctx.fillStyle = theme.selectionFill;
  ctx.beginPath();
  ctx.rect(selPx.x, selPx.y, selPx.w, selPx.h);
  ctx.rect(active.x, active.y, active.w, active.h);
  ctx.fill("evenodd");

  // References in the formula being edited.
  p.highlights.forEach((h, n) => {
    const color = REF_COLORS[n % REF_COLORS.length];
    const px = rect(h);
    ctx.fillStyle = color + "1a";
    ctx.fillRect(px.x, px.y, px.w, px.h);
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    ctx.strokeRect(px.x + 1, px.y + 1, px.w - 1, px.h - 1);
  });

  if (p.spill) {
    const px = rect(p.spill);
    ctx.strokeStyle = "#4472c4";
    ctx.lineWidth = 1;
    ctx.strokeRect(px.x + 0.5, px.y + 0.5, px.w - 1, px.h - 1);
  }

  // Circle Invalid Data: red ellipses around cells that break their rule.
  if (p.circles.length) {
    ctx.strokeStyle = theme.invalidCircle;
    ctx.lineWidth = 1.5;
    for (const cell of p.circles) {
      const px = rect(geo.merges.cellRect(cell));
      if (px.w <= 0 || px.h <= 0) continue;
      ctx.beginPath();
      ctx.ellipse(px.x + px.w / 2, px.y + px.h / 2, px.w / 2 + 3, px.h / 2 + 2, 0, 0, Math.PI * 2);
      ctx.stroke();
    }
  }

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
    const px = rect(p.fillPreview);
    ctx.setLineDash([4, 3]);
    ctx.strokeStyle = "#5f6368";
    ctx.lineWidth = 1;
    ctx.strokeRect(px.x + 0.5, px.y + 0.5, px.w - 1, px.h - 1);
    ctx.setLineDash([]);
  }
}

function paintHeaders(ctx: CanvasRenderingContext2D, p: PaintInput, panes: Panes) {
  const { width, height, geo, selection } = p;
  const { rows, cols } = geo;
  const sel = selectionRect(selection);
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(0, 0, width, HEADER_HEIGHT);
  ctx.fillRect(0, 0, HEADER_WIDTH, height);
  ctx.font = `13px ${theme.fontFamily}`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  const headerText = (inSel: boolean) => (inSel ? theme.headerActiveText : theme.headerText);
  // Excel shows the row numbers of a filtered list in blue.
  const f = p.filter;
  const filteredRow = (row: number) => !!f && f.columns.size > 0 && row > f.range.top && row <= f.range.bottom;

  for (const band of panes.colBands()) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(band.from, 0, band.to - band.from, HEADER_HEIGHT);
    ctx.clip();
    for (const col of cols.visibleLines(band.first, band.last)) {
      const x = HEADER_WIDTH + cols.x(col) - band.offset;
      const w = cols.width(col);
      const inSel = col >= sel.left && col <= sel.right;
      if (inSel) {
        ctx.fillStyle = theme.headerActiveBg;
        ctx.fillRect(x, 0, w, HEADER_HEIGHT);
        ctx.fillStyle = theme.headerActiveLine;
        ctx.fillRect(x, HEADER_HEIGHT - 2, w, 2);
      }
      ctx.fillStyle = headerText(inSel);
      ctx.fillText(colName(col), x + w / 2, HEADER_HEIGHT / 2);
      // A darker, double-width separator marks hidden columns, like Excel.
      const hiddenNext = col + 1 < MAX_COLS && cols.isHidden(col + 1);
      ctx.fillStyle = hiddenNext ? theme.hiddenMarker : theme.headerBorder;
      ctx.fillRect(Math.floor(x + w) - 1, 0, hiddenNext ? 2 : 1, HEADER_HEIGHT);
    }
    ctx.restore();
  }

  for (const band of panes.rowBands()) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, band.from, HEADER_WIDTH, band.to - band.from);
    ctx.clip();
    for (const row of rows.visibleLines(band.first, band.last)) {
      const y = HEADER_HEIGHT + rows.y(row) - band.offset;
      const h = rows.height(row);
      const inSel = row >= sel.top && row <= sel.bottom;
      if (inSel) {
        ctx.fillStyle = theme.headerActiveBg;
        ctx.fillRect(0, y, HEADER_WIDTH, h);
        ctx.fillStyle = theme.headerActiveLine;
        ctx.fillRect(HEADER_WIDTH - 2, y, 2, h);
      }
      ctx.fillStyle = filteredRow(row) && !inSel ? theme.filtered : headerText(inSel);
      ctx.fillText(String(row + 1), HEADER_WIDTH / 2, y + h / 2);
      const hiddenNext = row + 1 < MAX_ROWS && rows.isHidden(row + 1);
      ctx.fillStyle = hiddenNext ? theme.hiddenMarker : theme.headerBorder;
      ctx.fillRect(0, Math.floor(y + h) - 1, HEADER_WIDTH, hiddenNext ? 2 : 1);
    }
    ctx.restore();
  }
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

export { fontFor } from "./cellRenderer";
