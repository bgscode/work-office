// Viewport geometry with frozen panes. The grid is split into up to four
// panes: frozen rows/columns never scroll; the rest scroll as usual. Every
// position here is in logical (unzoomed) viewport pixels, headers included.
//
// With nothing frozen there is one pane and this is plain scrolling.

import { HEADER_HEIGHT, HEADER_WIDTH } from "./geometry";
import type { AxisLayout, ColumnLayout, RowLayout } from "./layout";

export interface Freeze {
  rows: number;
  cols: number;
}

export const NO_FREEZE: Freeze = Object.freeze({ rows: 0, cols: 0 });

/** One strip of the viewport along an axis: the frozen part or the scrolling part. */
export interface Band {
  /** Viewport start and end of the strip. */
  from: number;
  to: number;
  /** Scroll offset applied to lines in this strip (0 for the frozen strip). */
  offset: number;
  /** Lines that may be visible in it (inclusive; hidden ones included). */
  first: number;
  last: number;
}

/** A rectangle of the viewport and how lines map into it. */
export interface Pane {
  cols: Band;
  rows: Band;
}

export class Panes {
  /** Size of the frozen area, from the top-left of the cells. */
  readonly frozenWidth: number;
  readonly frozenHeight: number;

  constructor(
    readonly rows: RowLayout,
    readonly cols: ColumnLayout,
    readonly freeze: Freeze,
    readonly scroll: { top: number; left: number },
    readonly view: { width: number; height: number },
  ) {
    this.frozenWidth = cols.x(freeze.cols);
    this.frozenHeight = rows.y(freeze.rows);
  }

  /** Viewport x of a column's left edge. */
  colX(col: number): number {
    return HEADER_WIDTH + this.cols.x(col) - (col < this.freeze.cols ? 0 : this.scroll.left);
  }

  /** Viewport y of a row's top edge. */
  rowY(row: number): number {
    return HEADER_HEIGHT + this.rows.y(row) - (row < this.freeze.rows ? 0 : this.scroll.top);
  }

  /** Column under viewport x (in whichever pane x falls). */
  colAt(x: number): number {
    return lineAt(this.cols, x - HEADER_WIDTH, this.freeze.cols, this.frozenWidth, this.scroll.left);
  }

  /** Row under viewport y. */
  rowAt(y: number): number {
    return lineAt(this.rows, y - HEADER_HEIGHT, this.freeze.rows, this.frozenHeight, this.scroll.top);
  }

  colBands(): Band[] {
    return bands(this.cols, HEADER_WIDTH, this.view.width, this.freeze.cols, this.frozenWidth, this.scroll.left);
  }

  rowBands(): Band[] {
    return bands(this.rows, HEADER_HEIGHT, this.view.height, this.freeze.rows, this.frozenHeight, this.scroll.top);
  }

  panes(): Pane[] {
    const out: Pane[] = [];
    for (const rows of this.rowBands()) for (const cols of this.colBands()) out.push({ rows, cols });
    return out;
  }

  /**
   * The scroll offset (logical px) that brings line `i` fully into view along
   * one axis, or null if it is already visible or frozen.
   */
  static revealOffset(layout: AxisLayout, i: number, frozenLines: number, frozenSize: number, scroll: number, viewSize: number): number | null {
    if (i < frozenLines) return null;
    const start = layout.pos(i);
    const end = start + layout.size(i);
    if (start < scroll + frozenSize) return Math.max(0, start - frozenSize);
    if (end > scroll + viewSize) return Math.max(0, Math.min(start - frozenSize, end - viewSize));
    return null;
  }
}

function lineAt(layout: AxisLayout, local: number, frozenLines: number, frozenSize: number, scroll: number): number {
  if (frozenLines > 0 && local < frozenSize) return Math.min(layout.indexAt(Math.max(0, local)), frozenLines - 1);
  const i = layout.indexAt(local + scroll);
  return frozenLines > 0 ? Math.max(i, layout.nextVisible(frozenLines, 1) ?? frozenLines) : i;
}

function bands(layout: AxisLayout, header: number, viewEnd: number, frozenLines: number, frozenSize: number, scroll: number): Band[] {
  const out: Band[] = [];
  const split = Math.min(viewEnd, header + frozenSize);
  if (frozenLines > 0 && frozenSize > 0) {
    out.push({ from: header, to: split, offset: 0, first: 0, last: frozenLines - 1 });
  }
  if (split < viewEnd) {
    const first = Math.max(frozenLines, layout.indexAt(frozenSize + scroll));
    const last = Math.max(first, layout.indexAt(scroll + viewEnd - header));
    out.push({ from: split, to: viewEnd, offset: scroll, first, last: Math.min(last + 1, layout.count - 1) });
  }
  return out;
}
