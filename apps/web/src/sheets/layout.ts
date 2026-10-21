// Positions of rows and columns when some have custom sizes or are hidden.
//
// Sizes are sparse: a default size plus a short sorted list of segments
// (runs of lines with another size; hidden lines have size 0). Positions are
// computed from the segments in O(log segments), and finding the line under
// the mouse is a binary search, so a sheet with 1,048,576 rows costs nothing
// extra and nothing is allocated per row.

import { MAX_COLS, MAX_ROWS, ROW_HEIGHT } from "./geometry";

export const DEFAULT_COL_WIDTH = 100;
export const DEFAULT_ROW_HEIGHT = ROW_HEIGHT;

/** [first, last, size] inclusive run of lines. */
export type Run = [number, number, number];

/** Sizes along one axis (rows or columns). */
export class AxisLayout {
  // Segments of non-default size, sorted and non-overlapping.
  private readonly starts: number[] = [];
  private readonly ends: number[] = [];
  private readonly sizes: number[] = [];
  /** Extra length (vs. the default size) of all segments before segment i. */
  private readonly extra: number[] = [];
  private readonly sized: Run[];
  private readonly hidden: Array<[number, number]>;

  /**
   * `sized`: runs with a custom size; `hidden`: [first, last] runs of hidden
   * lines (they win over a custom size). Both sorted or not, any overlap is fine.
   */
  constructor(
    readonly count: number,
    readonly defaultSize: number,
    sized: Run[] = [],
    hidden: Array<[number, number]> = [],
  ) {
    this.sized = disjointRuns(sized);
    this.hidden = disjointPairs(hidden);
    // Sweep over every boundary and take the size of each piece.
    const cuts = new Set<number>();
    for (const [a, b] of this.sized) cuts.add(a).add(b + 1);
    for (const [a, b] of this.hidden) cuts.add(a).add(b + 1);
    const points = [...cuts].filter((p) => p >= 0 && p <= count).sort((a, b) => a - b);
    let extra = 0;
    for (let k = 0; k + 1 < points.length; k++) {
      const first = points[k];
      const last = points[k + 1] - 1;
      const size = find(this.hidden, first) ? 0 : (find(this.sized, first)?.[2] ?? defaultSize);
      if (size === defaultSize) continue;
      const n = this.starts.length;
      if (n > 0 && this.ends[n - 1] === first - 1 && this.sizes[n - 1] === size) {
        this.ends[n - 1] = last;
      } else {
        this.starts.push(first);
        this.ends.push(last);
        this.sizes.push(size);
        this.extra.push(extra);
      }
      extra += (last - first + 1) * (size - defaultSize);
    }
  }

  /** Index of the last segment starting at or before line i, or -1. */
  private segment(i: number): number {
    let lo = 0;
    let hi = this.starts.length - 1;
    let found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (this.starts[mid] <= i) {
        found = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    return found;
  }

  /** Start of line i, measured from the start of line 0. */
  pos(i: number): number {
    const k = this.segment(i);
    if (k < 0) return i * this.defaultSize;
    const inside = Math.min(i, this.ends[k] + 1) - this.starts[k];
    return i * this.defaultSize + this.extra[k] + inside * (this.sizes[k] - this.defaultSize);
  }

  size(i: number): number {
    const k = this.segment(i);
    return k >= 0 && i <= this.ends[k] ? this.sizes[k] : this.defaultSize;
  }

  isHidden(i: number): boolean {
    return this.size(i) === 0;
  }

  /** The visible line at position p (clamped to the valid range). */
  indexAt(p: number): number {
    if (p <= 0) return this.nextVisible(0, 1) ?? 0;
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.pos(mid) <= p) lo = mid;
      else hi = mid - 1;
    }
    return this.isHidden(lo) ? (this.nextVisible(lo, -1) ?? lo) : lo;
  }

  /** First visible line at or after i (dir 1) or at or before i (dir -1), or null. */
  nextVisible(i: number, dir: 1 | -1): number | null {
    let at = i;
    while (at >= 0 && at < this.count) {
      const k = this.segment(at);
      if (k < 0 || at > this.ends[k] || this.sizes[k] !== 0) return at;
      at = dir > 0 ? this.ends[k] + 1 : this.starts[k] - 1;
    }
    return null;
  }

  /** Visible lines in first..=last, ascending. */
  visibleLines(first: number, last: number): number[] {
    const out: number[] = [];
    let i = this.nextVisible(Math.max(0, first), 1);
    const end = Math.min(last, this.count - 1);
    while (i !== null && i <= end) {
      out.push(i);
      i = this.nextVisible(i + 1, 1);
    }
    return out;
  }

  /** A copy with line `i` set to `size` (used to preview a drag-resize). */
  withSize(i: number, size: number): this {
    const Ctor = this.constructor as new (count: number, def: number, sized: Run[], hidden: Array<[number, number]>) => this;
    const hidden = size > 0 ? this.hidden.filter(([a, b]) => i < a || i > b) : [...this.hidden, [i, i] as [number, number]];
    return new Ctor(this.count, this.defaultSize, [...this.sized, [i, i, size]], hidden);
  }
}

/** Column positions in pixels. */
export class ColumnLayout extends AxisLayout {
  /** `custom` is the flat [col, px, col, px, ...] list from `columnWidths`; `hidden` is [first, last, ...]. */
  constructor(custom: ArrayLike<number> | Run[] = [], hidden: ArrayLike<number> | Array<[number, number]> = []) {
    super(MAX_COLS, DEFAULT_COL_WIDTH, toRuns(custom), toPairs(hidden));
  }

  /** Left edge of a column, measured from the left of column A. */
  x(col: number): number {
    return this.pos(col);
  }

  width(col: number): number {
    return this.size(col);
  }

  /** Column containing position `x` (clamped to the valid range). */
  colAt(x: number): number {
    return this.indexAt(x);
  }
}

/** Row positions in pixels. */
export class RowLayout extends AxisLayout {
  constructor(sized: Run[] = [], hidden: Array<[number, number]> = []) {
    super(MAX_ROWS, DEFAULT_ROW_HEIGHT, sized, hidden);
  }

  /** From the engine's `rowLayout`: [n, (first, last, px) × n, (first, last) × rest]. */
  static fromEngine(flat: ArrayLike<number>): RowLayout {
    const n = flat[0] ?? 0;
    const sized: Run[] = [];
    for (let k = 0; k < n; k++) sized.push([flat[1 + 3 * k], flat[2 + 3 * k], flat[3 + 3 * k]]);
    const hidden: Array<[number, number]> = [];
    for (let i = 1 + 3 * n; i + 1 < flat.length; i += 2) hidden.push([flat[i], flat[i + 1]]);
    return new RowLayout(sized, hidden);
  }

  /** Top edge of a row, measured from the top of row 1. */
  y(row: number): number {
    return this.pos(row);
  }

  height(row: number): number {
    return this.size(row);
  }

  rowAt(y: number): number {
    return this.indexAt(y);
  }
}

function toRuns(custom: ArrayLike<number> | Run[]): Run[] {
  if (Array.isArray(custom) && Array.isArray(custom[0])) return custom as Run[];
  const flat = custom as ArrayLike<number>;
  const out: Run[] = [];
  for (let i = 0; i + 1 < flat.length; i += 2) out.push([flat[i], flat[i], flat[i + 1]]);
  return out;
}

function toPairs(hidden: ArrayLike<number> | Array<[number, number]>): Array<[number, number]> {
  if (Array.isArray(hidden) && Array.isArray(hidden[0])) return hidden as Array<[number, number]>;
  const flat = hidden as ArrayLike<number>;
  const out: Array<[number, number]> = [];
  for (let i = 0; i + 1 < flat.length; i += 2) out.push([flat[i], flat[i + 1]]);
  return out;
}

/** The run containing line i in sorted, disjoint runs (binary search), or null. */
function find<T extends [number, number, ...number[]]>(runs: T[], i: number): T | null {
  let lo = 0;
  let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (runs[mid][1] < i) lo = mid + 1;
    else if (runs[mid][0] > i) hi = mid - 1;
    else return runs[mid];
  }
  return null;
}

/** Sorted, disjoint runs; where runs overlap, later ones win. */
function disjointRuns(runs: Run[]): Run[] {
  let out: Run[] = [];
  let sorted = true;
  for (let k = 1; k < runs.length && sorted; k++) sorted = runs[k][0] > runs[k - 1][1];
  if (sorted) return runs;
  for (const [a, b, size] of runs) {
    const next: Run[] = [];
    for (const [s, e, v] of out) {
      if (e < a || s > b) next.push([s, e, v]);
      else {
        if (s < a) next.push([s, a - 1, v]);
        if (e > b) next.push([b + 1, e, v]);
      }
    }
    next.push([a, b, size]);
    out = next.sort((x, y) => x[0] - y[0]);
  }
  return out;
}

/** Sorted, merged [first, last] pairs. */
function disjointPairs(pairs: Array<[number, number]>): Array<[number, number]> {
  const sorted = [...pairs].sort((x, y) => x[0] - y[0]);
  const out: Array<[number, number]> = [];
  for (const [a, b] of sorted) {
    const last = out[out.length - 1];
    if (last && a <= last[1] + 1) last[1] = Math.max(last[1], b);
    else out.push([a, b]);
  }
  return out;
}
