import { describe, expect, it } from "vitest";
import { groupCommands, layoutGroup, type RibbonCommand } from "@office/ui";
import { autoSumRange, insertFunctionText, searchFunctions, sliderToZoom, statusMode, stepZoom, zoomToSlider } from "./chrome";

describe("statusMode", () => {
  it("follows Excel's Ready / Enter / Edit / Point", () => {
    expect(statusMode(null, false)).toBe("Ready");
    expect(statusMode({ draft: "1", source: "cell", mode: "enter", caret: 1, sheet: 0, cell: { row: 0, col: 0 } }, false)).toBe("Enter");
    expect(statusMode({ draft: "1", source: "cell", mode: "edit", caret: 1, sheet: 0, cell: { row: 0, col: 0 } }, false)).toBe("Edit");
    expect(statusMode({ draft: "=", source: "cell", mode: "enter", caret: 1, sheet: 0, cell: { row: 0, col: 0 } }, true)).toBe("Point");
  });
});

describe("zoom slider", () => {
  it("puts 100% in the middle", () => {
    expect(zoomToSlider(0.5)).toBe(0);
    expect(zoomToSlider(1)).toBe(50);
    expect(zoomToSlider(2)).toBe(100);
    expect(sliderToZoom(0)).toBe(0.5);
    expect(sliderToZoom(50)).toBe(1);
    expect(sliderToZoom(51)).toBe(1);
    expect(sliderToZoom(100)).toBe(2);
    expect(sliderToZoom(75)).toBe(1.5);
  });

  it("round-trips whole percents", () => {
    for (const z of [0.5, 0.75, 0.9, 1.1, 1.25, 1.5, 2]) expect(sliderToZoom(zoomToSlider(z))).toBe(z);
  });

  it("steps by 10%", () => {
    expect(stepZoom(1, 1)).toBe(1.1);
    expect(stepZoom(1, -1)).toBe(0.9);
    expect(stepZoom(1.25, 1)).toBe(1.3);
    expect(stepZoom(1.25, -1)).toBe(1.2);
    expect(stepZoom(2, 1)).toBe(2);
    expect(stepZoom(0.5, -1)).toBe(0.5);
  });
});

describe("insertFunctionText", () => {
  it("starts a formula when not in one", () => {
    expect(insertFunctionText(null, 0, "SUM")).toEqual({ draft: "=SUM(", caret: 5 });
    expect(insertFunctionText("hello", 2, "SUM")).toEqual({ draft: "=SUM(", caret: 5 });
  });

  it("inserts at the caret inside a formula", () => {
    expect(insertFunctionText("=A1+", 4, "MAX")).toEqual({ draft: "=A1+MAX(", caret: 8 });
    expect(insertFunctionText("=1+2", 1, "ABS")).toEqual({ draft: "=ABS(1+2", caret: 5 });
  });
});

describe("searchFunctions", () => {
  const all = [
    { name: "SUM", summary: "Adds numbers." },
    { name: "SUMIF", summary: "Adds numbers that meet a condition." },
    { name: "DSUM", summary: "Database sum." },
    { name: "AVERAGE", summary: "Arithmetic mean of numbers." },
  ];
  it("ranks name prefix, then name contains, then description", () => {
    expect(searchFunctions(all, "sum").map((f) => f.name)).toEqual(["SUM", "SUMIF", "DSUM"]);
    expect(searchFunctions(all, "mean").map((f) => f.name)).toEqual(["AVERAGE"]);
    expect(searchFunctions(all, "")).toHaveLength(4);
  });
});

describe("ribbon layout", () => {
  const cmd = (id: string, extra: Partial<RibbonCommand> = {}): RibbonCommand => ({ id, label: id, ...extra });

  it("groups commands in order of first appearance, ungrouped too", () => {
    const groups = groupCommands([cmd("a", { group: "Font" }), cmd("b"), cmd("c", { group: "Clipboard" }), cmd("d", { group: "Font" })]);
    expect(groups.map((g) => [g.name, g.commands.map((c) => c.id)])).toEqual([
      ["Font", ["a", "d"]],
      ["", ["b"]],
      ["Clipboard", ["c"]],
    ]);
  });

  it("stacks small buttons in threes and splits compact rows", () => {
    const segs = layoutGroup([
      cmd("paste", { size: "large" }),
      cmd("cut", { size: "small" }),
      cmd("copy", { size: "small" }),
      cmd("painter", { size: "small" }),
      cmd("more", { size: "small" }),
      cmd("b"),
      cmd("i"),
      cmd("u", { newRow: true }),
    ]);
    expect(segs.map((s) => s.kind)).toEqual(["large", "stack", "stack", "rows"]);
    expect(segs[3].kind === "rows" && segs[3].rows.map((r) => r.map((c) => c.id))).toEqual([["b", "i"], ["u"]]);
    const auto = layoutGroup(["a", "b", "c", "d", "e"].map((id) => cmd(id)));
    expect(auto[0].kind === "rows" && auto[0].rows.map((r) => r.length)).toEqual([3, 2]);
  });
});

describe("autoSumRange", () => {
  // Numbers in A1:A3 and B5:D5.
  const numbers = new Set(["0,0", "1,0", "2,0", "4,1", "4,2", "4,3"]);
  const isNumber = (r: number, c: number) => numbers.has(`${r},${c}`);
  it("adds up the numbers above, else to the left", () => {
    expect(autoSumRange(3, 0, isNumber)).toEqual({ top: 0, left: 0, bottom: 2, right: 0 });
    expect(autoSumRange(4, 4, isNumber)).toEqual({ top: 4, left: 1, bottom: 4, right: 3 });
    expect(autoSumRange(9, 9, isNumber)).toBeNull();
  });
});
