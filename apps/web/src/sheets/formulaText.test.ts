import { describe, expect, it } from "vitest";
import { canInsertRef, closeParens, cycleReference, enclosingFunction, findRefs, refsOnSheet, wordAtCaret } from "./formulaText";

describe("formulaText", () => {
  it("finds references but not function names or strings", () => {
    const refs = findRefs('=SUM(A1:B3)+$C$4+LOG10(D5)&"E6"');
    expect(refs.map((r) => [r.top, r.left, r.bottom, r.right])).toEqual([
      [0, 0, 2, 1],
      [3, 2, 3, 2],
      [4, 3, 4, 3],
    ]);
    expect(findRefs("A1+B2")).toEqual([]);
  });

  it("finds references to other sheets, whole columns and rows", () => {
    const refs = findRefs("=Sheet2!A1+'My Data'!B:C+SUM(3:5)+'It''s'!$A$1");
    expect(refs.map((r) => [r.sheet, r.top, r.left, r.bottom, r.right])).toEqual([
      ["Sheet2", 0, 0, 0, 0],
      ["My Data", 0, 1, 1_048_575, 2],
      [undefined, 2, 0, 4, 16_383],
      ["It's", 0, 0, 0, 0],
    ]);
    expect(findRefs("=LOG10(2)+SUM(1)")).toEqual([]);
    const onData = refsOnSheet(findRefs("=A1+Data!B2"), "Data", "Sheet1");
    expect(onData.map((r) => r.sheet)).toEqual(["Data"]);
  });

  it("cycles absolute references with F4", () => {
    expect(cycleReference("=A1+B2", 3)).toEqual({ formula: "=$A$1+B2", caret: 5 });
    expect(cycleReference("=$A$1", 5)?.formula).toBe("=A$1");
    expect(cycleReference("=A$1", 4)?.formula).toBe("=$A1");
    expect(cycleReference("=$A1", 4)?.formula).toBe("=A1");
    expect(cycleReference("=Data!A1:B2", 11)?.formula).toBe("=Data!$A$1:$B$2");
    expect(cycleReference("=SUM(", 5)).toBeNull();
  });

  it("finds the word being typed", () => {
    expect(wordAtCaret("=SUM(1)+vlo", 11)).toEqual({ start: 8, text: "vlo" });
    expect(wordAtCaret('="vlo', 5)).toBeNull();
    expect(wordAtCaret("=A1+", 4)).toBeNull();
  });

  it("knows the enclosing function", () => {
    expect(enclosingFunction("=IF(A1>0, SUM(B1, ", 18)).toBe("SUM");
    expect(enclosingFunction("=IF(A1>0, SUM(B1), ", 19)).toBe("IF");
    expect(enclosingFunction("=1+2", 4)).toBeNull();
  });

  it("decides where references can be inserted", () => {
    expect(canInsertRef("=", 1)).toBe(true);
    expect(canInsertRef("=SUM(", 5)).toBe(true);
    expect(canInsertRef("=A1+ ", 5)).toBe(true);
    expect(canInsertRef("=A1", 3)).toBe(false);
    expect(canInsertRef('="a,', 4)).toBe(false);
    expect(canInsertRef("hello,", 6)).toBe(false);
  });

  it("closes open parentheses", () => {
    expect(closeParens("=SUM(IF(A1>0,1")).toBe("=SUM(IF(A1>0,1))");
    expect(closeParens('=LEN("(")')).toBe('=LEN("(")');
  });
});
