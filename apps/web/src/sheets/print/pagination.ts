// Splits a print range into pages the way Excel does: column bands across,
// row bands down, pages numbered down then over. Sizes are CSS pixels at 100%
// (96 per inch). Rows/columns to repeat (print titles) are added to every page
// that doesn't already start with them, and take space from it.

export type Span = [number, number];

export interface Band {
  start: number;
  end: number;
  /** Title rows/columns repeated before this band, or null. */
  titles: Span | null;
}

export interface PaginateInput {
  /** Inclusive sheet coordinates of what to print. */
  rows: Span;
  cols: Span;
  rowHeight: (row: number) => number;
  colWidth: (col: number) => number;
  titleRows: Span | null;
  titleCols: Span | null;
  /** Printable area of one page (paper minus margins), px. */
  availWidth: number;
  availHeight: number;
  /** Size of printed row/column headings (0 when off), px at 100%. */
  headingWidth: number;
  headingHeight: number;
  scaling: { percent: number } | { fitWidth: number; fitHeight: number };
}

export interface Pagination {
  scale: number;
  rowBands: Band[];
  colBands: Band[];
}

/** Excel's smallest scale. */
const MIN_SCALE = 0.1;

function sum(size: (i: number) => number, a: number, b: number): number {
  let total = 0;
  for (let i = a; i <= b; i++) total += size(i);
  return total;
}

/** Greedy bands: as many lines as fit, at least one per page. */
export function bands(span: Span, size: (i: number) => number, avail: number, titles: Span | null, fixed: number, scale: number): Band[] {
  const out: Band[] = [];
  let start = span[0];
  while (start <= span[1]) {
    const repeat: Span | null = titles && titles[0] < start ? [titles[0], Math.min(titles[1], start - 1)] : null;
    let used = (fixed + (repeat ? sum(size, repeat[0], repeat[1]) : 0) + size(start)) * scale;
    let end = start;
    while (end < span[1]) {
      const next = used + size(end + 1) * scale;
      if (next > avail + 0.01) break;
      used = next;
      end++;
    }
    out.push({ start, end, titles: repeat });
    start = end + 1;
  }
  return out;
}

export function paginate(p: PaginateInput): Pagination {
  const across = (scale: number) => bands(p.cols, p.colWidth, p.availWidth, p.titleCols, p.headingWidth, scale);
  const down = (scale: number) => bands(p.rows, p.rowHeight, p.availHeight, p.titleRows, p.headingHeight, scale);

  if ("percent" in p.scaling) {
    const scale = Math.min(4, Math.max(MIN_SCALE, p.scaling.percent / 100));
    return { scale, rowBands: down(scale), colBands: across(scale) };
  }

  const { fitWidth, fitHeight } = p.scaling;
  // First guess from the totals, then shrink until the page counts fit (whole
  // rows/columns per page can need a little more).
  const guess = (pages: number, total: number, titles: number, fixed: number, avail: number) =>
    pages > 0 ? (pages * avail) / (total + pages * fixed + (pages - 1) * titles) : Infinity;
  const titleSize = (t: Span | null, size: (i: number) => number) => (t ? sum(size, t[0], t[1]) : 0);
  const totalW = sum(p.colWidth, p.cols[0], p.cols[1]);
  const totalH = sum(p.rowHeight, p.rows[0], p.rows[1]);
  let scale = Math.min(
    1,
    guess(fitWidth, totalW, titleSize(p.titleCols, p.colWidth), p.headingWidth, p.availWidth),
    guess(fitHeight, totalH, titleSize(p.titleRows, p.rowHeight), p.headingHeight, p.availHeight),
  );
  scale = Math.max(MIN_SCALE, scale);
  for (let i = 0; i < 200; i++) {
    const cols = fitWidth > 0 ? across(scale) : null;
    const rows = fitHeight > 0 ? down(scale) : null;
    const fits = (!cols || cols.length <= fitWidth) && (!rows || rows.length <= fitHeight);
    if (fits || scale <= MIN_SCALE) break;
    scale = Math.max(MIN_SCALE, scale * 0.98);
  }
  // Excel rounds the scale down to a whole percent.
  scale = Math.max(MIN_SCALE, Math.floor(scale * 100 + 1e-9) / 100);
  return { scale, rowBands: down(scale), colBands: across(scale) };
}

/** Pages in Excel's default order: down, then over. */
export function pageOrder(pg: Pagination): { row: Band; col: Band }[] {
  const out: { row: Band; col: Band }[] = [];
  for (const col of pg.colBands) for (const row of pg.rowBands) out.push({ row, col });
  return out;
}
