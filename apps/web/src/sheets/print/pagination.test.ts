import { describe, expect, it } from "vitest";
import { bands, pageOrder, paginate, type PaginateInput } from "./pagination";

const base: PaginateInput = {
  rows: [0, 99],
  cols: [0, 9],
  rowHeight: () => 20,
  colWidth: () => 100,
  titleRows: null,
  titleCols: null,
  availWidth: 450,
  availHeight: 500,
  headingWidth: 0,
  headingHeight: 0,
  scaling: { percent: 100 },
};

describe("bands", () => {
  it("fits as many lines as possible and at least one", () => {
    expect(bands([0, 9], () => 100, 250, null, 0, 1).map((b) => [b.start, b.end])).toEqual([
      [0, 1],
      [2, 3],
      [4, 5],
      [6, 7],
      [8, 9],
    ]);
    // A line wider than the page still gets its own page.
    expect(bands([0, 1], () => 900, 250, null, 0, 1).map((b) => [b.start, b.end])).toEqual([
      [0, 0],
      [1, 1],
    ]);
  });

  it("repeats titles on later pages and makes room for them", () => {
    const b = bands([0, 9], () => 10, 50, [0, 0], 0, 1);
    expect(b[0]).toEqual({ start: 0, end: 4, titles: null });
    expect(b[1]).toEqual({ start: 5, end: 8, titles: [0, 0] });
    expect(b[2]).toEqual({ start: 9, end: 9, titles: [0, 0] });
  });

  it("repeats titles above the print area on the first page too", () => {
    expect(bands([5, 6], () => 10, 50, [0, 1], 0, 1)[0].titles).toEqual([0, 1]);
  });
});

describe("paginate", () => {
  it("splits at actual size", () => {
    const pg = paginate(base);
    expect(pg.scale).toBe(1);
    expect(pg.colBands.length).toBe(3); // 4 + 4 + 2 columns
    expect(pg.rowBands.length).toBe(4); // 25 rows per page
    const order = pageOrder(pg);
    expect(order.length).toBe(12);
    // Down, then over.
    expect(order[1].row.start).toBe(25);
    expect(order[1].col.start).toBe(0);
    expect(order[4].col.start).toBe(4);
  });

  it("fits the sheet on one page", () => {
    const pg = paginate({ ...base, scaling: { fitWidth: 1, fitHeight: 1 } });
    expect(pg.colBands.length).toBe(1);
    expect(pg.rowBands.length).toBe(1);
    expect(pg.scale).toBeLessThanOrEqual(0.25);
    expect(pg.scale).toBeGreaterThan(0.2);
  });

  it("fits all columns on one page and lets rows flow", () => {
    const pg = paginate({ ...base, scaling: { fitWidth: 1, fitHeight: 0 } });
    expect(pg.colBands.length).toBe(1);
    expect(pg.scale).toBe(0.45);
    expect(pg.rowBands.length).toBe(2); // 55 rows per page at 45%
  });

  it("never shrinks below 10%", () => {
    const pg = paginate({ ...base, rows: [0, 99999], scaling: { fitWidth: 1, fitHeight: 1 } });
    expect(pg.scale).toBe(0.1);
    expect(pg.rowBands.length).toBeGreaterThan(1);
  });

  it("does not enlarge small sheets", () => {
    const pg = paginate({ ...base, rows: [0, 1], cols: [0, 0], scaling: { fitWidth: 1, fitHeight: 1 } });
    expect(pg.scale).toBe(1);
  });

  it("counts headings", () => {
    const pg = paginate({ ...base, headingWidth: 60 });
    expect(pg.colBands[0].end).toBe(2); // 60 + 3 * 100 <= 450 < 60 + 4 * 100
  });
});
