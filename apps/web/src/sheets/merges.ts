// Merged cells and keyboard navigation over merged and hidden cells.
// Pure functions of the sheet layout, so they are easy to test
// (merges.test.ts). The engine owns the merges; this is a read-only index.

import { MAX_COLS, MAX_ROWS, rangeName, selectionRect, single, type Cell, type Rect, type Selection } from "./geometry";
import type { AxisLayout } from "./layout";

const sameRect = (a: Rect, b: Rect) => a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
const overlaps = (a: Rect, b: Rect) => a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;

export class MergeIndex {
  readonly list: Rect[];

  /** `flat` is the engine's `merges`: [top, left, bottom, right, ...]. */
  constructor(flat: ArrayLike<number> = []) {
    this.list = [];
    for (let i = 0; i + 3 < flat.length; i += 4) {
      this.list.push({ top: flat[i], left: flat[i + 1], bottom: flat[i + 2], right: flat[i + 3] });
    }
  }

  get isEmpty(): boolean {
    return this.list.length === 0;
  }

  /** The merged area containing a cell. */
  at(row: number, col: number): Rect | null {
    for (const m of this.list) if (row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) return m;
    return null;
  }

  /** Merged areas overlapping a rectangle. */
  within(r: Rect): Rect[] {
    return this.list.filter((m) => overlaps(m, r));
  }

  /** Grow a rectangle until no merged area sticks out of it (Excel's selection rule). */
  expand(r: Rect): Rect {
    let out = r;
    for (;;) {
      let grown = out;
      for (const m of this.within(out)) {
        grown = {
          top: Math.min(grown.top, m.top),
          left: Math.min(grown.left, m.left),
          bottom: Math.max(grown.bottom, m.bottom),
          right: Math.max(grown.right, m.right),
        };
      }
      if (sameRect(grown, out)) return out;
      out = grown;
    }
  }

  /**
   * A selection as Excel would have it: the active cell is the top-left of
   * its merged area, and the range covers every merged area it touches.
   */
  normalize(s: Selection): Selection {
    if (this.isEmpty) return s.span ? { anchor: s.anchor, focus: s.focus } : s;
    const m = this.at(s.anchor.row, s.anchor.col);
    const anchor = m ? { row: m.top, col: m.left } : s.anchor;
    const plain = selectionRect({ anchor, focus: s.focus });
    const base = m
      ? { top: Math.min(plain.top, m.top), left: Math.min(plain.left, m.left), bottom: Math.max(plain.bottom, m.bottom), right: Math.max(plain.right, m.right) }
      : plain;
    const span = this.expand(base);
    return sameRect(span, plain) ? { anchor, focus: s.focus } : { anchor, focus: s.focus, span };
  }

  /** True when the selection is exactly one merged area (it then behaves like one cell). */
  isSingle(s: Selection): boolean {
    const m = this.at(s.anchor.row, s.anchor.col);
    return !!m && sameRect(m, selectionRect(s));
  }

  /** Name Box text: "B2" for a merged cell, like Excel, otherwise the range. */
  label(s: Selection): string {
    return rangeName(this.isSingle(s) ? single(s.anchor) : s);
  }

  /** The rectangle a cell occupies: its merged area or just itself. */
  cellRect(c: Cell): Rect {
    return this.at(c.row, c.col) ?? { top: c.row, left: c.col, bottom: c.row, right: c.col };
  }
}

export interface NavLayout {
  rows: AxisLayout;
  cols: AxisLayout;
  merges: MergeIndex;
}

/** Move `d` lines from `i`, then onto a visible line (searching onward first, then back). */
function stepVisible(layout: AxisLayout, i: number, d: number, max: number): number {
  if (d === 0) return i;
  const target = Math.min(Math.max(i + d, 0), max - 1);
  const dir = d > 0 ? 1 : -1;
  const v = layout.nextVisible(target, dir) ?? layout.nextVisible(target, dir > 0 ? -1 : 1);
  // Nowhere visible to go in that direction: stay put.
  if (v === null || (dir > 0 ? v < i : v > i)) return i;
  return v;
}

/**
 * Arrow keys, Enter, Tab, Page Up/Down: move the active cell (or with
 * `extend` the far corner) by (dRow, dCol), skipping hidden rows and columns
 * and stepping over merged areas as one cell, like Excel.
 */
export function navigate(s: Selection, dRow: number, dCol: number, extend: boolean, g: NavLayout): Selection {
  const from = extend ? s.focus : s.anchor;
  // Leave a merged area from its far edge in the direction of travel.
  const m = g.merges.at(from.row, from.col);
  let row = from.row;
  let col = from.col;
  if (m) {
    if (dRow > 0) row = m.bottom;
    if (dRow < 0) row = m.top;
    if (dCol > 0) col = m.right;
    if (dCol < 0) col = m.left;
  }
  const to = { row: stepVisible(g.rows, row, dRow, MAX_ROWS), col: stepVisible(g.cols, col, dCol, MAX_COLS) };
  // Didn't move (edge of the sheet): keep the selection as it was.
  if (to.row === row && to.col === col && (dRow !== 0 || dCol !== 0)) {
    return extend ? g.merges.normalize(s) : g.merges.normalize(single(from));
  }
  return g.merges.normalize(extend ? { anchor: s.anchor, focus: to } : single(to));
}
