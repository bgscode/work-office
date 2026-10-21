// Grid sizes and pure selection math. No React or engine imports here, so it
// is easy to unit-test (see geometry.test.ts).

export const ROW_HEIGHT = 20;

export const HEADER_WIDTH = 50;
export const HEADER_HEIGHT = 20;
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export interface Cell {
  row: number;
  col: number;
}

/** `anchor` is the active cell; `focus` is the opposite corner of the range. */
export interface Selection {
  anchor: Cell;
  focus: Cell;
}

export interface Rect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export function colName(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function cellName(c: Cell): string {
  return `${colName(c.col)}${c.row + 1}`;
}

export function rangeName(s: Selection): string {
  const r = selectionRect(s);
  if (r.top === r.bottom && r.left === r.right) return cellName(s.anchor);
  return `${cellName({ row: r.top, col: r.left })}:${cellName({ row: r.bottom, col: r.right })}`;
}

export function selectionRect(s: Selection): Rect {
  return {
    top: Math.min(s.anchor.row, s.focus.row),
    left: Math.min(s.anchor.col, s.focus.col),
    bottom: Math.max(s.anchor.row, s.focus.row),
    right: Math.max(s.anchor.col, s.focus.col),
  };
}

export function clampCell(c: Cell): Cell {
  return {
    row: Math.min(Math.max(c.row, 0), MAX_ROWS - 1),
    col: Math.min(Math.max(c.col, 0), MAX_COLS - 1),
  };
}

export function single(c: Cell): Selection {
  return { anchor: c, focus: c };
}

/** Arrow-key behaviour: move the active cell, or with Shift grow the range. */
export function moveSelection(s: Selection, dRow: number, dCol: number, extend: boolean): Selection {
  if (extend) {
    return { anchor: s.anchor, focus: clampCell({ row: s.focus.row + dRow, col: s.focus.col + dCol }) };
  }
  return single(clampCell({ row: s.anchor.row + dRow, col: s.anchor.col + dCol }));
}

/** Intersect a rect with the area that holds data, or null if they don't overlap. */
export function clampToExtent(r: Rect, rows: number, cols: number): Rect | null {
  const bottom = Math.min(r.bottom, rows - 1);
  const right = Math.min(r.right, cols - 1);
  if (bottom < r.top || right < r.left) return null;
  return { top: r.top, left: r.left, bottom, right };
}
