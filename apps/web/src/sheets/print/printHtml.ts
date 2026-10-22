// Turns sheets into a printable HTML document: one fixed-size <section> per
// page (paper size, margins, header/footer), each holding a <table> of that
// page's cells with their formatting. The same HTML drives the print preview
// and the real print (a hidden iframe), so what you see is what prints; the
// browser's "Save as PDF" printer gives PDF export.
//
// Pure: cell data comes through `SheetSource`, text measuring through
// `measure`, so it is unit-tested without a browser.

import type { CellStyle } from "@office/engine";
import { renderHeaderFooter, type HeaderFooterContext } from "./headerFooter";
import { pageInches, type PageSetup } from "./pageSetup";
import { pageOrder, paginate, type Band, type Span } from "./pagination";

/** Value kinds from `kindBlock` (see `ValueKind` in `@office/engine`). */
const KIND_NUMBER = 1;
const KIND_TEXT = 2;

/**
 * Cell style as printed. The optional fields are read when the engine
 * provides them (fonts, vertical alignment, wrap, borders from formatting work);
 * unknown fields are ignored.
 */
export type PrintStyle = CellStyle & {
  /** Font family name. */
  font?: string;
  /** Font size in points. */
  size?: number;
  valign?: "top" | "middle" | "center" | "bottom";
  wrap?: 1 | boolean;
  /** CSS border shorthands per side, e.g. "1px solid #000". */
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
};

export interface Rect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface CellBlock {
  /** Row-major for the requested rectangle. */
  texts: string[];
  kinds: ArrayLike<number>;
  styles: PrintStyle[];
}

export interface SheetSource {
  name: string;
  setup: PageSetup;
  /** What to print (inclusive), or null when the sheet has nothing to print. */
  range: Rect | null;
  colWidth(col: number): number;
  rowHeight(row: number): number;
  block(r: Rect): CellBlock;
  /** Merged areas (printed as one cell, clipped to the page). */
  merges?: Rect[];
  /** Charts and pictures as image URLs, placed in sheet pixels (from the top-left of A1). */
  images?: PrintImage[];
  /** Left edge of a column / top edge of a row in sheet pixels; needed for images. */
  colX?(col: number): number;
  rowY?(row: number): number;
}

export interface PrintImage {
  src: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface PrintJob {
  /** File name, for `&F` and the document title. */
  title: string;
  sheets: SheetSource[];
  date: Date;
  /** Width in px of text drawn in a style (canvas measureText in the app). */
  measure(text: string, style: PrintStyle): number;
}

export interface PageInfo {
  sheet: number;
  row: Band;
  col: Band;
  scale: number;
}

export const PX_PER_INCH = 96;
/** Excel's default: Calibri 11pt. */
const BASE_FONT = "Calibri, Carlito, 'Segoe UI', Arial, sans-serif";
const BASE_SIZE_PX = 15;
const CELL_PAD = 4;
const HEADING_HEIGHT = 20;

function headingWidth(lastRow: number): number {
  return Math.max(3, String(lastRow + 1).length) * 8 + 10;
}

function spanOf(s: string | null, parse: (a: string) => number): Span | null {
  if (!s) return null;
  const [a, b = a] = s.replace(/\$/g, "").split(":");
  const x = parse(a);
  const y = parse(b);
  return Number.isNaN(x) || Number.isNaN(y) ? null : [Math.min(x, y), Math.max(x, y)];
}

/** "1:2" -> [0, 1]. */
export function rowSpan(s: string | null): Span | null {
  return spanOf(s, (a) => (/^\d+$/.test(a) ? Number(a) - 1 : NaN));
}

/** "A:B" -> [0, 1]. */
export function colSpan(s: string | null): Span | null {
  return spanOf(s, (a) => {
    if (!/^[A-Za-z]{1,3}$/.test(a)) return NaN;
    let n = 0;
    for (const ch of a.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
    return n - 1;
  });
}

export function colLetters(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Split every sheet of the job into pages. */
export function layoutJob(job: PrintJob): PageInfo[] {
  const pages: PageInfo[] = [];
  job.sheets.forEach((src, sheet) => {
    const r = src.range;
    if (!r) return;
    const setup = src.setup;
    const paper = pageInches(setup);
    const m = setup.margins;
    const pg = paginate({
      rows: [r.top, r.bottom],
      cols: [r.left, r.right],
      rowHeight: (i) => src.rowHeight(i),
      colWidth: (i) => src.colWidth(i),
      titleRows: rowSpan(setup.titleRows),
      titleCols: colSpan(setup.titleCols),
      availWidth: Math.max(0.5, paper.width - m.left - m.right) * PX_PER_INCH,
      availHeight: Math.max(0.5, paper.height - m.top - m.bottom) * PX_PER_INCH,
      headingWidth: setup.headings ? headingWidth(r.bottom) : 0,
      headingHeight: setup.headings ? HEADING_HEIGHT : 0,
      scaling: setup.fit ? { fitWidth: setup.fit[0], fitHeight: setup.fit[1] } : { percent: setup.scale },
    });
    for (const { row, col } of pageOrder(pg)) pages.push({ sheet, row, col, scale: pg.scale });
  });
  return pages;
}

export function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Lines in sheet order: repeated titles first, then the band. */
function linesOf(band: Band): number[] {
  const out: number[] = [];
  if (band.titles) for (let i = band.titles[0]; i <= band.titles[1]; i++) out.push(i);
  for (let i = band.start; i <= band.end; i++) out.push(i);
  return out;
}

/** Contiguous runs of a sorted list, as spans. */
function runs(lines: number[]): Span[] {
  const out: Span[] = [];
  for (const l of lines) {
    const last = out[out.length - 1];
    if (last && last[1] === l - 1) last[1] = l;
    else out.push([l, l]);
  }
  return out;
}

/** Cell lookup for a page, fetching one block per row run x column run. */
function pageCells(src: SheetSource, rows: number[], cols: number[]) {
  const blocks: { r: Span; c: Span; data: CellBlock }[] = [];
  for (const r of runs(rows)) {
    for (const c of runs(cols)) {
      blocks.push({ r, c, data: src.block({ top: r[0], left: c[0], bottom: r[1], right: c[1] }) });
    }
  }
  return (row: number, col: number) => {
    const b = blocks.find((x) => row >= x.r[0] && row <= x.r[1] && col >= x.c[0] && col <= x.c[1])!;
    const i = (row - b.r[0]) * (b.c[1] - b.c[0] + 1) + (col - b.c[0]);
    return { text: b.data.texts[i] ?? "", kind: b.data.kinds[i] ?? 0, style: b.data.styles[i] ?? {} };
  };
}

function fontCss(style: PrintStyle): string {
  const parts: string[] = [];
  if (style.b) parts.push("font-weight:bold");
  if (style.i) parts.push("font-style:italic");
  const deco = [style.u ? "underline" : "", style.s ? "line-through" : ""].filter(Boolean).join(" ");
  if (deco) parts.push(`text-decoration:${deco}`);
  if (style.font) parts.push(`font-family:${cssFontFamily(style.font)}`);
  if (style.size) parts.push(`font-size:${style.size}pt`);
  return parts.join(";");
}

function cssFontFamily(name: string): string {
  return `"${name.replace(/["\\<>]/g, "")}",${BASE_FONT}`;
}

const SAFE_COLOR = /^#[0-9a-fA-F]{3,8}$/;
const SAFE_BORDER = /^[\w\s.#%(),-]+$/;

function cellCss(style: PrintStyle, align: string, wrap: boolean): string {
  const parts = [fontCss(style)];
  if (style.color && SAFE_COLOR.test(style.color)) parts.push(`color:${style.color}`);
  if (style.fill && SAFE_COLOR.test(style.fill)) parts.push(`background:${style.fill}`);
  if (align !== "left") parts.push(`text-align:${align}`);
  if (style.valign) parts.push(`vertical-align:${style.valign === "center" ? "middle" : style.valign}`);
  if (wrap) parts.push("white-space:pre-wrap;word-break:break-word");
  for (const [side, value] of [
    ["top", style.borderTop],
    ["right", style.borderRight],
    ["bottom", style.borderBottom],
    ["left", style.borderLeft],
  ] as const) {
    if (value && SAFE_BORDER.test(value)) parts.push(`border-${side}:${value}`);
  }
  return parts.filter(Boolean).join(";");
}

function hfHtml(code: string, ctx: HeaderFooterContext, cls: string, style: string): string {
  if (!code) return "";
  const s = renderHeaderFooter(code, ctx);
  return (
    `<div class="hf ${cls}" style="${style}">` +
    `<div class="l">${escapeHtml(s.left)}</div><div class="c">${escapeHtml(s.center)}</div><div class="r">${escapeHtml(s.right)}</div>` +
    `</div>`
  );
}

function pageTable(job: PrintJob, page: PageInfo): string {
  const src = job.sheets[page.sheet];
  const setup = src.setup;
  const rows = linesOf(page.row).filter((r) => src.rowHeight(r) > 0);
  const cols = linesOf(page.col).filter((c) => src.colWidth(c) > 0);
  const cell = pageCells(src, rows, cols);
  // Merged areas: the first printed cell of each gets row/colspan over the printed lines it covers.
  const rowIndex = new Map(rows.map((r, i) => [r, i]));
  const colIndex = new Map(cols.map((c, i) => [c, i]));
  const covered = new Set<string>();
  const spans = new Map<string, [number, number]>();
  for (const m of src.merges ?? []) {
    const mr = rows.filter((r) => r >= m.top && r <= m.bottom);
    const mc = cols.filter((c) => c >= m.left && c <= m.right);
    if (!mr.length || !mc.length || (mr.length === 1 && mc.length === 1)) continue;
    // Only contiguous printed runs merge; titles + band are kept apart.
    const rEnd = mr.findIndex((r, i) => i > 0 && rowIndex.get(r)! !== rowIndex.get(mr[i - 1])! + 1);
    const cEnd = mc.findIndex((c, i) => i > 0 && colIndex.get(c)! !== colIndex.get(mc[i - 1])! + 1);
    const rs = mr.slice(0, rEnd < 0 ? mr.length : rEnd);
    const cs = mc.slice(0, cEnd < 0 ? mc.length : cEnd);
    spans.set(`${rs[0]},${cs[0]}`, [rs.length, cs.length]);
    for (const r of rs) for (const c of cs) if (r !== rs[0] || c !== cs[0]) covered.add(`${r},${c}`);
  }
  const grid = setup.gridlines;
  const headW = setup.headings ? headingWidth(src.range?.bottom ?? 0) : 0;

  let html = `<table class="${grid ? "grid" : ""}">`;
  html += "<colgroup>";
  if (headW) html += `<col style="width:${headW}px">`;
  for (const c of cols) html += `<col style="width:${src.colWidth(c)}px">`;
  html += "</colgroup>";
  if (setup.headings) {
    html += `<tr style="height:${HEADING_HEIGHT}px"><th class="corner"></th>`;
    for (const c of cols) html += `<th>${colLetters(c)}</th>`;
    html += "</tr>";
  }
  for (const r of rows) {
    html += `<tr style="height:${src.rowHeight(r)}px">`;
    if (setup.headings) html += `<th>${r + 1}</th>`;
    for (let i = 0; i < cols.length; i++) {
      const c = cols[i];
      if (covered.has(`${r},${c}`)) continue;
      const merged = spans.get(`${r},${c}`);
      const { text: raw, kind, style } = cell(r, c);
      const wrap = !!style.wrap;
      let text = wrap ? raw : raw.replace(/\n/g, " ");
      const align = style.align ?? (kind === KIND_NUMBER ? "right" : kind === KIND_TEXT || !text ? "left" : "center");
      const width = src.colWidth(c);
      let span = 1;
      if (text && !wrap && !merged) {
        const textWidth = job.measure(text, style);
        if (kind === KIND_NUMBER && textWidth > width - 2 * CELL_PAD) {
          // Numbers that don't fit print as ### like Excel.
          const hash = Math.max(1, job.measure("#", style));
          text = "#".repeat(Math.max(1, Math.floor((width - 2 * CELL_PAD) / hash)));
        } else if (align === "left" && textWidth > width - 2 * CELL_PAD) {
          // Text spills over empty neighbours on the same page.
          let room = width;
          while (i + span < cols.length && cols[i + span] === c + span && room < textWidth + 2 * CELL_PAD) {
            const next = cell(r, c + span);
            if (next.text || next.style.fill) break;
            room += src.colWidth(c + span);
            span++;
          }
        }
      }
      const css = cellCss(style, align, wrap);
      const colspan = merged ? merged[1] : span;
      const rowspan = merged ? ` rowspan="${merged[0]}"` : "";
      html += `<td${colspan > 1 ? ` colspan="${colspan}"` : ""}${rowspan}${css ? ` style="${escapeHtml(css)}"` : ""}>${escapeHtml(text)}</td>`;
      i += span - 1;
    }
    html += "</tr>";
  }
  html += "</table>";
  return `<div class="sheet" style="zoom:${page.scale}">${html}${pageImages(src, page, headW)}</div>`;
}

/** Charts and pictures over this page's band, clipped to it. */
function pageImages(src: SheetSource, page: PageInfo, headW: number): string {
  if (!src.images?.length || !src.colX || !src.rowY) return "";
  const colX = src.colX;
  const rowY = src.rowY;
  const titleW = page.col.titles ? colX(page.col.titles[1] + 1) - colX(page.col.titles[0]) : 0;
  const titleH = page.row.titles ? rowY(page.row.titles[1] + 1) - rowY(page.row.titles[0]) : 0;
  const x0 = colX(page.col.start);
  const y0 = rowY(page.row.start);
  const w = colX(page.col.end + 1) - x0;
  const h = rowY(page.row.end + 1) - y0;
  const headH = src.setup.headings ? HEADING_HEIGHT : 0;
  const shown = src.images.filter((im) => im.x < x0 + w && im.x + im.w > x0 && im.y < y0 + h && im.y + im.h > y0);
  if (!shown.length) return "";
  let out = `<div class="images" style="left:${headW + titleW}px;top:${headH + titleH}px;width:${w}px;height:${h}px">`;
  for (const im of shown) {
    if (!/^(data:image\/|blob:)/.test(im.src)) continue;
    out += `<img src="${escapeHtml(im.src)}" style="left:${im.x - x0}px;top:${im.y - y0}px;width:${im.w}px;height:${im.h}px">`;
  }
  return out + "</div>";
}

export interface HtmlOptions {
  /** Pages to include (indexes into `pages`); default all. */
  include?: number[];
  /** Preview: no page gaps or shadows, for showing a single page scaled. */
  preview?: boolean;
}

/** The full HTML document for printing (or previewing) `pages` of `job`. */
export function buildPrintHtml(job: PrintJob, pages: PageInfo[], opts: HtmlOptions = {}): string {
  const include = opts.include ?? pages.map((_, i) => i);
  let pageRules = "";
  job.sheets.forEach((src, i) => {
    const size = pageInches(src.setup);
    const width = +size.width.toFixed(4);
    const height = +size.height.toFixed(4);
    pageRules += `@page s${i}{size:${width}in ${height}in;margin:0}.s${i}{page:s${i};width:${width}in;height:${height}in}`;
  });
  const css =
    `@page{margin:0}` +
    pageRules +
    `html,body{margin:0;padding:0;background:#fff}` +
    `body{font:${BASE_SIZE_PX}px ${BASE_FONT};color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact}` +
    `.page{position:relative;overflow:hidden;background:#fff;break-after:page;page-break-after:always}` +
    `.page:last-child{break-after:auto;page-break-after:auto}` +
    `.area{position:absolute;display:flex;overflow:hidden}` +
    `.hf{position:absolute;display:flex;font-size:11px;white-space:pre}` +
    `.hf>div{flex:1;overflow:hidden}.hf .c{text-align:center}.hf .r{text-align:right}` +
    `.sheet{position:relative;flex:none}` +
    `.images{position:absolute;overflow:hidden}.images img{position:absolute}` +
    `table{border-collapse:collapse;table-layout:fixed;width:max-content;flex:none}` +
    `td,th{padding:0 ${CELL_PAD}px;overflow:hidden;white-space:pre;vertical-align:bottom;line-height:1.25;border:0}` +
    `table.grid td{border:1px solid #bfbfbf}` +
    `th{font-weight:normal;font-size:11px;text-align:center;vertical-align:middle;border:1px solid #bfbfbf;color:#000}` +
    (opts.preview ? "" : `@media screen{html,body{background:#e6e6e6}.page{margin:16px auto;box-shadow:0 1px 4px rgba(0,0,0,.3)}}`);

  const ctxBase = { pages: pages.length, file: job.title, date: job.date };
  let body = "";
  for (const index of include) {
    const page = pages[index];
    if (!page) continue;
    const src = job.sheets[page.sheet];
    const s = src.setup;
    const m = s.margins;
    const ctx: HeaderFooterContext = { ...ctxBase, page: index + 1, sheet: src.name };
    const justify = s.centerH ? "center" : "flex-start";
    const alignItems = s.centerV ? "center" : "flex-start";
    body +=
      `<section class="page s${page.sheet}" data-page="${index + 1}">` +
      hfHtml(s.header, ctx, "header", `top:${m.header}in;left:${m.left}in;right:${m.right}in`) +
      `<div class="area" style="top:${m.top}in;left:${m.left}in;right:${m.right}in;bottom:${m.bottom}in;justify-content:${justify};align-items:${alignItems}">` +
      pageTable(job, page) +
      `</div>` +
      hfHtml(s.footer, ctx, "footer", `bottom:${m.footer}in;left:${m.left}in;right:${m.right}in`) +
      `</section>`;
  }
  return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(job.title)}</title><style>${css}</style></head><body>${body}</body></html>`;
}
