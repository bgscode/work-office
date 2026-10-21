import { describe, expect, it } from "vitest";
import { parseTsv, toTsv } from "./clipboard";

describe("clipboard", () => {
  it("round-trips tricky cells", () => {
    const rows = [
      ["a", "b\tc"],
      ['say "hi"', "line1\nline2"],
    ];
    expect(parseTsv(toTsv(rows))).toEqual(rows);
  });

  it("ignores the trailing newline Excel adds", () => {
    expect(parseTsv("1\t2\r\n3\t4\r\n")).toEqual([
      ["1", "2"],
      ["3", "4"],
    ]);
  });
});
