import { describe, expect, it } from "vitest";
import { adjacentNote, linkDestination } from "./annotations";

describe("linkDestination", () => {
  it("opens web, e-mail and workbook links like Excel", () => {
    expect(linkDestination("https://example.com")).toEqual({ kind: "web", href: "https://example.com" });
    expect(linkDestination("www.example.com")).toEqual({ kind: "web", href: "https://www.example.com" });
    expect(linkDestination("asha@example.com")).toEqual({ kind: "web", href: "mailto:asha@example.com" });
    expect(linkDestination("#'Sales 2025'!B4")).toEqual({ kind: "place", ref: "'Sales 2025'!B4" });
  });

  it("never runs scripts", () => {
    expect(linkDestination("javascript:alert(1)")).toBeNull();
    expect(linkDestination(" JavaScript:alert(1)")).toBeNull();
    expect(linkDestination("data:text/html,x")).toBeNull();
    expect(linkDestination("")).toBeNull();
  });
});

describe("adjacentNote", () => {
  // Notes in B2, D2 and A5 (row, col pairs in row order).
  const cells = [1, 1, 1, 3, 4, 0];
  it("goes to the next or previous note, wrapping around", () => {
    expect(adjacentNote(cells, 1, 1, 1)).toEqual({ row: 1, col: 3 });
    expect(adjacentNote(cells, 4, 0, 1)).toEqual({ row: 1, col: 1 });
    expect(adjacentNote(cells, 1, 1, -1)).toEqual({ row: 4, col: 0 });
    expect(adjacentNote([], 0, 0, 1)).toBeNull();
  });
});
