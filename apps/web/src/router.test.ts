import { describe, expect, it } from "vitest";
import { parseRoute } from "./router";

describe("parseRoute", () => {
  it("reads pages and document ids", () => {
    expect(parseRoute("")).toEqual({ page: "home" });
    expect(parseRoute("#/")).toEqual({ page: "home" });
    expect(parseRoute("#/sheets/abc-123")).toEqual({ page: "sheets", docId: "abc-123" });
    expect(parseRoute("#/sheets")).toEqual({ page: "sheets", docId: undefined });
    expect(parseRoute("#/docs")).toEqual({ page: "docs", docId: undefined });
    expect(parseRoute("#/nonsense")).toEqual({ page: "home" });
  });
});
