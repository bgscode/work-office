import { describe, expect, it } from "vitest";
import { ColumnLayout } from "./layout";

describe("ColumnLayout", () => {
  it("handles custom widths", () => {
    const layout = new ColumnLayout([1, 40, 3, 250]);
    expect(layout.x(2)).toBe(140);
    expect(layout.width(3)).toBe(250);
    expect(layout.colAt(0)).toBe(0);
    expect(layout.colAt(139)).toBe(1);
    expect(layout.colAt(140)).toBe(2);
    expect(layout.colAt(400)).toBe(3);
  });
});
