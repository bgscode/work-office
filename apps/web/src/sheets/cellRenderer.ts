// Draws one cell's content (text with font, colour, alignment, wrap,
// underline/strikethrough) and its borders on the grid canvas. paint.ts
// decides *which* cells are visible and where they are; this file decides how
// a cell looks inside its rectangle, so it works with any row height or merged
// area. The measuring helpers are shared with column/row autofit.

import { ValueKind, type BorderLine, type CellBorders, type CellStyle } from "@office/engine";

/** Excel's default font for new workbooks. */
export const DEFAULT_FONT = "Calibri";
/** Default font size in points. */
export const DEFAULT_FONT_SIZE = 11;
/** Space between the cell edge and its text, in logical pixels. */
export const CELL_PADDING = 3;

/** Fonts offered in Home → Font, in Excel's order of popularity. */
export const FONT_CHOICES = [
  "Calibri",
  "Aptos Narrow",
  "Arial",
  "Cambria",
  "Courier New",
  "Georgia",
  "Mangal",
  "Nirmala UI",
  "Segoe UI",
  "Tahoma",
  "Times New Roman",
  "Verdana",
];

/** Sizes offered in Home → Font Size (points). */
export const FONT_SIZE_CHOICES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** Fallbacks with similar metrics, so files look right where the font isn't installed. */
const FALLBACKS: Record<string, string> = {
  calibri: `Carlito, system-ui, -apple-system, "Segoe UI", sans-serif`,
  "aptos narrow": `Aptos, "Arial Narrow", system-ui, -apple-system, sans-serif`,
  arial: `"Liberation Sans", Helvetica, sans-serif`,
  cambria: `Caladea, Georgia, serif`,
  "courier new": `"Liberation Mono", Courier, monospace`,
  georgia: `serif`,
  mangal: `"Nirmala UI", "Kohinoor Devanagari", "Noto Sans Devanagari", sans-serif`,
  "nirmala ui": `Mangal, "Kohinoor Devanagari", "Noto Sans Devanagari", sans-serif`,
  "segoe ui": `system-ui, -apple-system, sans-serif`,
  tahoma: `Verdana, sans-serif`,
  "times new roman": `"Liberation Serif", Times, serif`,
  verdana: `sans-serif`,
};

/** CSS font-family list for an Excel font name. */
export function fontFamilyCss(name: string | undefined): string {
  const family = name || DEFAULT_FONT;
  const quoted = `"${family.replace(/["\\]/g, "")}"`;
  return `${quoted}, ${FALLBACKS[family.toLowerCase()] ?? "system-ui, -apple-system, sans-serif"}`;
}

const installed = new Map<string, boolean>();

/** Whether a font is installed, by comparing its text width with generic fonts. */
export function isFontInstalled(name: string): boolean {
  let known = installed.get(name);
  if (known !== undefined) return known;
  known = true;
  if (typeof document !== "undefined") {
    const ctx = document.createElement("canvas").getContext("2d");
    if (ctx) {
      const sample = "mmmmmmmmmlli10WQ@";
      known = ["monospace", "serif"].some((generic) => {
        ctx.font = `72px ${generic}`;
        const base = ctx.measureText(sample).width;
        ctx.font = `72px "${name}", ${generic}`;
        return ctx.measureText(sample).width !== base;
      });
    }
  }
  installed.set(name, known);
  return known;
}

/**
 * Calibri is small for its point size; when it falls back to a system font the
 * text would look a size too large, so it is scaled to the same visual size.
 */
function fallbackScale(name: string | undefined): number {
  const family = (name || DEFAULT_FONT).toLowerCase();
  if (family !== "calibri" && family !== "aptos narrow") return 1;
  return isFontInstalled(name || DEFAULT_FONT) ? 1 : 0.89;
}

/** Font size in logical pixels (1 pt = 4/3 px, as in Excel at 100% zoom). */
export function fontSizePx(style: CellStyle): number {
  return ((style.sz ?? DEFAULT_FONT_SIZE) * 4) / 3 * fallbackScale(style.font);
}

/** Canvas `font` shorthand for a cell style. */
export function fontFor(style: CellStyle): string {
  const px = Math.round(fontSizePx(style) * 100) / 100;
  return `${style.i ? "italic " : ""}${style.b ? "bold " : ""}${px}px ${fontFamilyCss(style.font)}`;
}

/** Height of one line of text in this style (logical pixels). */
export function lineHeightFor(style: CellStyle): number {
  return Math.ceil(fontSizePx(style) * 1.3);
}

/**
 * Break text into the lines a wrapped cell shows: at line breaks (Alt+Enter),
 * then between words, and inside words that are wider than the cell.
 */
export function wrapLines(text: string, maxWidth: number, measure: (s: string) => number): string[] {
  const out: string[] = [];
  for (const paragraph of text.split("\n")) {
    if (measure(paragraph) <= maxWidth) {
      out.push(paragraph);
      continue;
    }
    let line = "";
    for (const word of paragraph.split(/(?<= )/)) {
      const candidate = line + word;
      if (measure(candidate.trimEnd()) <= maxWidth) {
        line = candidate;
        continue;
      }
      if (line) out.push(line.trimEnd());
      line = "";
      // A word longer than the cell is broken between characters.
      let rest = word;
      while (measure(rest.trimEnd()) > maxWidth && rest.length > 1) {
        let n = 1;
        while (n < rest.length && measure(rest.slice(0, n + 1)) <= maxWidth) n++;
        out.push(rest.slice(0, n));
        rest = rest.slice(n);
      }
      line = rest;
    }
    out.push(line.trimEnd());
  }
  return out;
}

/** The lines shown for `text` in a cell `width` pixels wide (one line unless wrapped). */
export function cellLines(ctx: CanvasRenderingContext2D, text: string, style: CellStyle, width: number): string[] {
  if (!style.wrap) return [text.includes("\n") ? text.replace(/\n/g, " ") : text];
  ctx.font = fontFor(style);
  return wrapLines(text, Math.max(1, width - 2 * CELL_PADDING), (s) => ctx.measureText(s).width);
}

/**
 * Height a cell needs to show all its text at `width` (for row autofit): the
 * wrapped line count times the line height, plus padding.
 */
export function cellTextHeight(ctx: CanvasRenderingContext2D, text: string, style: CellStyle, width: number): number {
  return cellLines(ctx, text, style, width).length * lineHeightFor(style) + 4;
}

export interface CellRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface DrawCellOptions {
  /** Default text colour. */
  textColor: string;
  /** Colour for error values (#DIV/0! …). */
  errorColor: string;
  /** Horizontal clip when text overflows into empty neighbours (defaults to the cell). */
  clipLeft?: number;
  clipRight?: number;
}

type HAlign = "left" | "center" | "right";

/** Excel's "General" alignment: numbers right, text left, booleans and errors centred. */
export function horizontalAlign(kind: number, style: CellStyle): HAlign {
  return style.align ?? (kind === ValueKind.Number ? "right" : kind === ValueKind.Text ? "left" : "center");
}

/**
 * How far one-line text sticks out of its cell on each side, in pixels.
 * Wrapped text and numbers never overflow (numbers show ### instead).
 */
export function textOverflow(
  ctx: CanvasRenderingContext2D,
  rect: CellRect,
  text: string,
  kind: number,
  style: CellStyle,
): { left: number; right: number } {
  if (!text || style.wrap || kind === ValueKind.Number) return { left: 0, right: 0 };
  ctx.font = fontFor(style);
  const width = ctx.measureText(cellLines(ctx, text, style, rect.w)[0]).width;
  const extra = width - (rect.w - 2 * CELL_PADDING);
  if (extra <= 0) return { left: 0, right: 0 };
  switch (horizontalAlign(kind, style)) {
    case "left":
      return { left: 0, right: extra };
    case "right":
      return { left: extra, right: 0 };
    case "center":
      return { left: extra / 2, right: extra / 2 };
  }
}

/**
 * How far overflowing text may spill: walks over neighbouring columns while
 * `isFree(col)` says they are empty, until the text fits. Returns the clip edges.
 */
export function spillClip(
  rect: CellRect,
  overflow: { left: number; right: number },
  col: number,
  isFree: (col: number) => boolean,
  widthOf: (col: number) => number,
): { left: number; right: number } {
  let right = rect.x + rect.w;
  for (let c = col + 1; overflow.right > 0 && right < rect.x + rect.w + overflow.right && isFree(c); c++) right += widthOf(c);
  let left = rect.x;
  for (let c = col - 1; overflow.left > 0 && c >= 0 && left > rect.x - overflow.left && isFree(c); c--) left -= widthOf(c);
  return { left, right };
}

/** Draw a cell's text inside `rect` (logical pixels). The caller paints the fill and grid lines. */
export function drawCell(
  ctx: CanvasRenderingContext2D,
  rect: CellRect,
  text: string,
  kind: number,
  style: CellStyle,
  opts: DrawCellOptions,
) {
  if (!text) return;
  const { x, y, w, h } = rect;
  const pad = CELL_PADDING;
  ctx.font = fontFor(style);
  const align = horizontalAlign(kind, style);
  let lines = cellLines(ctx, text, style, w);

  // Numbers that don't fit show ### like Excel, rather than a misleading partial number.
  if (kind === ValueKind.Number && ctx.measureText(lines[0]).width > w - 2 * pad) {
    lines = ["#".repeat(Math.max(1, Math.floor((w - 2 * pad) / ctx.measureText("#").width)))];
  }

  const px = fontSizePx(style);
  const lineH = lineHeightFor(style);
  const blockH = lines.length * lineH;
  const top = style.va === "top" ? y + 2 : style.va === "center" ? y + (h - blockH) / 2 : y + h - blockH - 2;

  const clipLeft = opts.clipLeft ?? x;
  const clipRight = opts.clipRight ?? x + w;
  if (clipLeft < x || clipRight > x + w) {
    // Overflowing text covers the grid lines of the empty cells it spills into.
    ctx.fillStyle = "#ffffff";
    if (clipLeft < x) ctx.fillRect(clipLeft + 1, y + 1, x - clipLeft, h - 1);
    if (clipRight > x + w) ctx.fillRect(x + w, y + 1, clipRight - (x + w), h - 1);
  }

  ctx.save();
  ctx.beginPath();
  ctx.rect(clipLeft + 1, y + 1, clipRight - clipLeft - 2, h - 1);
  ctx.clip();
  ctx.fillStyle = kind === ValueKind.Error ? opts.errorColor : (style.color ?? opts.textColor);
  ctx.strokeStyle = ctx.fillStyle;
  ctx.textBaseline = "middle";
  ctx.lineWidth = Math.max(1, Math.round(px / 15));
  lines.forEach((line, i) => {
    const mid = Math.round(top + i * lineH + lineH / 2);
    if (mid - lineH / 2 > y + h || mid + lineH / 2 < y) return;
    const width = ctx.measureText(line).width;
    const tx = align === "right" ? x + w - pad - width : align === "center" ? x + (w - width) / 2 : x + pad;
    ctx.fillText(line, tx, mid);
    if ((style.u || style.s) && width > 0) {
      ctx.beginPath();
      if (style.u) {
        const uy = Math.round(mid + px * 0.5) + 0.5;
        ctx.moveTo(tx, uy);
        ctx.lineTo(tx + width, uy);
      }
      if (style.s) {
        const sy = Math.round(mid) + 0.5;
        ctx.moveTo(tx, sy);
        ctx.lineTo(tx + width, sy);
      }
      ctx.stroke();
    }
  });
  ctx.restore();
}

/** Thickness in pixels and dash pattern of each Excel border style. */
const LINE: Record<BorderLine["style"], { width: number; dash: number[] }> = {
  thin: { width: 1, dash: [] },
  medium: { width: 2, dash: [] },
  thick: { width: 3, dash: [] },
  double: { width: 3, dash: [] },
  hair: { width: 1, dash: [1, 1] },
  dotted: { width: 1, dash: [1, 2] },
  dashed: { width: 1, dash: [3, 1] },
  mediumDashed: { width: 2, dash: [9, 3] },
  dashDot: { width: 1, dash: [9, 3, 3, 3] },
  mediumDashDot: { width: 2, dash: [9, 3, 3, 3] },
  dashDotDot: { width: 1, dash: [9, 3, 3, 3, 3, 3] },
  mediumDashDotDot: { width: 2, dash: [9, 3, 3, 3, 3, 3] },
  slantDashDot: { width: 2, dash: [11, 1, 5, 1] },
};

/** Line width in pixels of a border style (for hit areas and previews). */
export function borderWidth(style: BorderLine["style"]): number {
  return LINE[style].width;
}

/**
 * Draw the borders of one cell on its edges. Lines are centred on the grid
 * line, like Excel, and extend into the corners so thick outlines join up.
 */
export function drawBorders(ctx: CanvasRenderingContext2D, rect: CellRect, border: CellBorders | undefined, background = "#ffffff") {
  if (!border) return;
  const x0 = Math.floor(rect.x);
  const y0 = Math.floor(rect.y);
  const x1 = Math.floor(rect.x + rect.w);
  const y1 = Math.floor(rect.y + rect.h);
  const reach = (line: BorderLine | undefined) => (line ? Math.floor(LINE[line.style].width / 2) : 0);
  const horizontal = (line: BorderLine | undefined, at: number) => {
    if (!line) return;
    const from = x0 - reach(border.left);
    const to = x1 + reach(border.right) + 1;
    stroke(ctx, line, background, (offset) => [from, at + offset, to, at + offset]);
  };
  const vertical = (line: BorderLine | undefined, at: number) => {
    if (!line) return;
    const from = y0 - reach(border.top);
    const to = y1 + reach(border.bottom) + 1;
    stroke(ctx, line, background, (offset) => [at + offset, from, at + offset, to]);
  };
  horizontal(border.top, y0);
  horizontal(border.bottom, y1);
  vertical(border.left, x0);
  vertical(border.right, x1);
}

/** Stroke one border line; `at(offset)` gives its endpoints shifted across the line. */
function stroke(ctx: CanvasRenderingContext2D, line: BorderLine, background: string, at: (offset: number) => number[]) {
  const spec = LINE[line.style];
  const color = line.color ?? "#000000";
  ctx.save();
  ctx.lineCap = "butt";
  ctx.setLineDash(spec.dash);
  const seg = (offset: number, width: number, style: string) => {
    const [ax, ay, bx, by] = at(offset);
    ctx.strokeStyle = style;
    ctx.lineWidth = width;
    ctx.beginPath();
    ctx.moveTo(ax, ay);
    ctx.lineTo(bx, by);
    ctx.stroke();
  };
  if (line.style === "double") {
    // Two thin lines with a gap that hides the grid line.
    seg(0.5, 1, background);
    seg(-0.5, 1, color);
    seg(1.5, 1, color);
  } else {
    // Odd widths sit on pixel centres; 2px lines cover the grid line and the pixel after it.
    seg(spec.width % 2 === 1 ? 0.5 : 1, spec.width, color);
  }
  ctx.restore();
}

/** Font settings for the in-cell editor so typing looks like the cell. */
export function editorFont(style: CellStyle): { fontFamily: string; fontSize: number; lineHeight: number; fontWeight?: string; fontStyle?: string } {
  return {
    fontFamily: fontFamilyCss(style.font),
    fontSize: fontSizePx(style),
    lineHeight: lineHeightFor(style),
    fontWeight: style.b ? "bold" : undefined,
    fontStyle: style.i ? "italic" : undefined,
  };
}
