import { describe, expect, it } from "vitest";
import { clampToExtent, colName, moveSelection, rangeName, single } from "./geometry";

describe("geometry", () => {
  it("names columns like Excel", () => {
    expect(colName(0)).toBe("A");
    expect(colName(25)).toBe("Z");
    expect(colName(26)).toBe("AA");
    expect(colName(16_383)).toBe("XFD");
  });

  it("moves and extends selections", () => {
    const s = single({ row: 0, col: 0 });
    expect(moveSelection(s, -1, 0, false)).toEqual(single({ row: 0, col: 0 }));
    const ext = moveSelection(moveSelection(s, 2, 0, true), 0, 1, true);
    expect(rangeName(ext)).toBe("A1:B3");
  });

  it("clamps ranges to the data extent", () => {
    expect(clampToExtent({ top: 0, left: 0, bottom: 1_000_000, right: 3 }, 10, 2)).toEqual({
      top: 0,
      left: 0,
      bottom: 9,
      right: 1,
    });
    expect(clampToExtent({ top: 20, left: 0, bottom: 30, right: 0 }, 10, 2)).toBeNull();
  });
});
