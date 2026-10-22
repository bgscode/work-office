import { describe, expect, it } from "vitest";
import { ColumnLayout, DEFAULT_COL_WIDTH } from "./layout";

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
