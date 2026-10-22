import { describe, expect, it } from "vitest";
import { ColumnLayout, DEFAULT_COL_WIDTH, RowLayout } from "./layout";

describe("ColumnLayout", () => {
  it("uses Excel's default column width", () => {
    expect(DEFAULT_COL_WIDTH).toBe(64);
    expect(new ColumnLayout().x(3)).toBe(192);
  });

  it("handles custom widths", () => {
    const layout = new ColumnLayout([1, 40, 3, 250]);
    expect(layout.x(2)).toBe(104);
    expect(layout.width(3)).toBe(250);
    expect(layout.colAt(0)).toBe(0);
    expect(layout.colAt(103)).toBe(1);
    expect(layout.colAt(104)).toBe(2);
    expect(layout.colAt(400)).toBe(3);
  });
});

describe("withSize (live preview while dragging a border)", () => {
  it("keeps the other columns' widths", () => {
    const layout = new ColumnLayout([1, 40, 3, 250]).withSize(2, 100);
    expect(layout).toBeInstanceOf(ColumnLayout);
    expect([layout.width(1), layout.width(2), layout.width(3)]).toEqual([40, 100, 250]);
  });

  it("works for rows, with hidden rows", () => {
    // This used to throw: RowLayout was built with ColumnLayout's arguments.
    const rows = new RowLayout([[0, 0, 40]], [[5, 6]]).withSize(2, 33);
    expect(rows).toBeInstanceOf(RowLayout);
    expect([rows.height(0), rows.height(2), rows.height(5)]).toEqual([40, 33, 0]);
    expect(rows.withSize(5, 25).height(5)).toBe(25);
  });
});
