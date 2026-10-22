import { describe, expect, it } from "vitest";
import { moveTarget, nextVisible, textOn } from "./sheetTabs";

describe("sheet tabs", () => {
  const tabs = [false, true, false, true].map((hidden) => ({ color: null, hidden }));

  it("skips hidden sheets when moving between tabs", () => {
    expect(nextVisible(tabs, 0, 1)).toBe(2);
    expect(nextVisible(tabs, 2, -1)).toBe(0);
    expect(nextVisible(tabs, 2, 1)).toBe(2);
  });

  it("turns a drop position into the moved sheet's index", () => {
    expect(moveTarget(0, 3)).toBe(2);
    expect(moveTarget(3, 0)).toBe(0);
    expect(moveTarget(1, 4)).toBe(3);
  });

  it("picks readable text on a tab colour", () => {
    expect(textOn("#FFC000")).toBe("#000000");
    expect(textOn("#002060")).toBe("#FFFFFF");
  });
});
