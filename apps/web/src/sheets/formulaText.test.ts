import { describe, expect, it } from "vitest";
import { canInsertRef, closeParens, enclosingFunction, findRefs, wordAtCaret } from "./formulaText";

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
