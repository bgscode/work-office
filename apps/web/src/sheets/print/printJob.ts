// Builds a PrintJob (printHtml.ts) from the open workbook: what to print on
// each sheet (selection, print area or the used range), row/column sizes from
// the grid's layout, cell text and formatting (with conditional formatting),
// merged cells, and charts/pictures as images.

import { parseStyle, type CellStyle, type SheetsDocument } from "@office/engine";
import { fontFor } from "../cellRenderer";
import { anchorRect } from "../charts/DrawingLayer";
import { drawChart, axisLengths } from "../charts/renderChart";
import { drawShape } from "../charts/renderShape";
import type { ChartView, DrawingInfo } from "../charts/types";
import { MAX_COLS, MAX_ROWS, type Rect } from "../geometry";
import { buildGeometry } from "../sheetGeometry";
import { readPageSetup } from "./pageSetup";
import { readTabs } from "../features/sheetTabs";
import { colSpan, rowSpan, type PrintImage, type PrintJob, type PrintStyle, type SheetSource } from "./printHtml";

export type PrintWhat = "active" | "workbook" | "selection";

const BORDER_CSS: Record<string, string> = {
  thin: "1px solid",
  hair: "1px dotted",
  dotted: "1px dotted",
  dashed: "1px dashed",
  medium: "2px solid",
  mediumDashed: "2px dashed",
  thick: "3px solid",
  double: "3px double",
};

/** Engine cell style -> printed style. */
export function printStyle(style: CellStyle): PrintStyle {
  const out: PrintStyle = { ...style };
  if (style.sz) out.size = style.sz;
  if (style.va) out.valign = style.va;
  const b = style.border;
  if (b) {
    const css = (line?: { style: string; color?: string }) => line && `${BORDER_CSS[line.style] ?? "1px dashed"} ${line.color ?? "#000"}`;
    out.borderTop = css(b.top);
    out.borderRight = css(b.right);
    out.borderBottom = css(b.bottom);
    out.borderLeft = css(b.left);
  }
  return out;
}

/** "A1:D20", "A:D" or "1:20" -> rectangle, or null. */
export function parseArea(area: string): Rect | null {
  const a = area.replace(/\$/g, "").trim();
  const cell = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/.exec(a);
  if (cell) {
    const c1 = colSpan(`${cell[1]}:${cell[3]}`)!;
    return { top: Math.min(+cell[2], +cell[4]) - 1, bottom: Math.max(+cell[2], +cell[4]) - 1, left: c1[0], right: c1[1] };
  }
  const cols = colSpan(a);
  if (cols) return { top: 0, bottom: MAX_ROWS - 1, left: cols[0], right: cols[1] };
  const rows = rowSpan(a);
  if (rows) return { top: rows[0], bottom: rows[1], left: 0, right: MAX_COLS - 1 };
  return null;
}

function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `data:${mime};base64,${btoa(bin)}`;
}

/** Draw at twice the resolution so printouts are sharp. */
function canvasImage(w: number, h: number, draw: (ctx: CanvasRenderingContext2D) => void): string | null {
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(2, Math.round(w * 2));
  canvas.height = Math.max(2, Math.round(h * 2));
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.scale(2, 2);
  draw(ctx);
  return canvas.toDataURL("image/png");
}

function chartImage(doc: SheetsDocument, sheet: number, id: number, w: number, h: number): string | null {
  const lengths = axisLengths(w, h);
  const json = doc.chartViewJson(sheet, id, lengths.value, lengths.x);
  if (!json) return null;
  const canvas = document.createElement("canvas");
  // Twice the resolution so printed charts are sharp.
  canvas.width = Math.round(w * 2);
  canvas.height = Math.round(h * 2);
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.scale(2, 2);
  drawChart(ctx, JSON.parse(json) as ChartView, w, h);
  return canvas.toDataURL("image/png");
}

function sheetSource(doc: SheetsDocument, sheet: number, selection: Rect | null): SheetSource {
  const setup = readPageSetup(doc, sheet);
  const geo = buildGeometry(doc, sheet);
  const [rows, cols] = doc.extent(sheet);
  const drawings = JSON.parse(doc.drawingsJson(sheet) || "[]") as DrawingInfo[];
  const placed = drawings.map((d) => ({ d, r: anchorRect(geo, d.anchor) }));

  // Used range: cells with data, grown to cover charts and pictures (as Excel does).
  let range: Rect | null = rows > 0 && cols > 0 ? { top: 0, left: 0, bottom: rows - 1, right: cols - 1 } : null;
  for (const { r } of placed) {
    const bottom = geo.rows.rowAt(r.y + r.h);
    const right = geo.cols.colAt(r.x + r.w);
    range = range
      ? { ...range, bottom: Math.max(range.bottom, bottom), right: Math.max(range.right, right) }
      : { top: 0, left: 0, bottom, right };
  }
  const clamp = (r: Rect): Rect | null => {
    if (!range) return null;
    const out = { top: r.top, left: r.left, bottom: Math.min(r.bottom, range.bottom), right: Math.min(r.right, range.right) };
    return out.bottom < out.top || out.right < out.left ? null : out;
  };
  const area = selection ?? (setup.printArea ? parseArea(setup.printArea) : null);
  const printRange = area ? clamp(area) : range;

  const merges: Rect[] = [];
  const flat = doc.merges(sheet);
  for (let i = 0; i + 3 < flat.length; i += 4) merges.push({ top: flat[i], left: flat[i + 1], bottom: flat[i + 2], right: flat[i + 3] });

  const images: PrintImage[] = [];
  for (const { d, r } of placed) {
    if (r.w <= 0 && r.h <= 0) continue;
    const src =
      d.kind === "chart"
        ? chartImage(doc, sheet, d.id, r.w, r.h)
        : d.kind === "shape"
          ? canvasImage(r.w, r.h, (ctx) => drawShape(ctx, d.shape, r.w, r.h))
          : bytesToDataUrl(doc.pictureData(sheet, d.id), d.mime);
    if (src) images.push({ src, ...r });
  }

  return {
    name: doc.sheetNames()[sheet],
    setup,
    range: printRange,
    colWidth: (c) => geo.cols.width(c),
    rowHeight: (r) => geo.rows.height(r),
    colX: (c) => geo.cols.x(c),
    rowY: (r) => geo.rows.y(r),
    merges,
    images,
    block(r) {
      const h = r.bottom - r.top + 1;
      const w = r.right - r.left + 1;
      const styles = doc.styleBlock(sheet, r.top, r.left, h, w).map((json) => printStyle(parseStyle(json)));
      // Conditional formatting prints too (fills, colours, bold...).
      const cf = doc.cfBlock(sheet, r.top, r.left, h, w);
      cf.forEach((json, i) => {
        if (!json) return;
        const { bar: _bar, barColor: _barColor, ...rest } = JSON.parse(json) as CellStyle & { bar?: number; barColor?: string };
        styles[i] = { ...styles[i], ...rest };
      });
      return { texts: doc.displayBlock(sheet, r.top, r.left, h, w), kinds: doc.kindBlock(sheet, r.top, r.left, h, w), styles };
    },
  };
}

/** The print job for `what` on the active sheet, the whole workbook, or the selection. */
export function buildPrintJob(doc: SheetsDocument, activeSheet: number, what: PrintWhat, selection: Rect, title: string): PrintJob {
  const sheets =
    what === "workbook"
      ? // Hidden sheets don't print, as in Excel.
        readTabs(doc).flatMap((tab, i) => (tab.hidden ? [] : [sheetSource(doc, i, null)]))
      : [sheetSource(doc, activeSheet, what === "selection" ? selection : null)];
  const ctx = document.createElement("canvas").getContext("2d")!;
  return {
    title,
    sheets,
    date: new Date(),
    measure(text, style) {
      ctx.font = fontFor(style);
      return ctx.measureText(text).width;
    },
  };
}
