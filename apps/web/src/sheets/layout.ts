// Column positions when columns can have different widths. Offsets are
// prefix sums, so finding the column under the mouse is a binary search.

import { MAX_COLS } from "./geometry";

export const DEFAULT_COL_WIDTH = 64;

export class ColumnLayout {
  private readonly offsets: Float64Array;

  /** `custom` is the flat [col, px, col, px, ...] list from `columnWidths`. */
  constructor(custom: ArrayLike<number> = []) {
    const widths = new Float64Array(MAX_COLS).fill(DEFAULT_COL_WIDTH);
    for (let i = 0; i + 1 < custom.length; i += 2) widths[custom[i]] = custom[i + 1];
    this.offsets = new Float64Array(MAX_COLS + 1);
    for (let c = 0; c < MAX_COLS; c++) this.offsets[c + 1] = this.offsets[c] + widths[c];
  }

  /** Left edge of a column, measured from the left of column A. */
  x(col: number): number {
    return this.offsets[col];
  }

  width(col: number): number {
    return this.offsets[col + 1] - this.offsets[col];
  }

  /** Column containing position `x` (clamped to the valid range). */
  colAt(x: number): number {
    if (x <= 0) return 0;
    let lo = 0;
    let hi = MAX_COLS - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offsets[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}
