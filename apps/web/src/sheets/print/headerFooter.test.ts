import { describe, expect, it } from "vitest";
import { renderHeaderFooter } from "./headerFooter";

const ctx = { page: 2, pages: 5, sheet: "Sales", file: "Book1", date: new Date(2025, 7, 15, 14, 30) };

describe("renderHeaderFooter", () => {
  it("fills page fields into the centre by default", () => {
    expect(renderHeaderFooter("Page &P of &N", ctx)).toEqual({ left: "", center: "Page 2 of 5", right: "" });
  });

  it("splits sections and names", () => {
    expect(renderHeaderFooter("&L&F&C&A&RPage &P", ctx)).toEqual({ left: "Book1", center: "Sales", right: "Page 2" });
  });

  it("drops font codes and keeps literal ampersands", () => {
    expect(renderHeaderFooter('&"Arial,Bold"&14&BR&&D &KFF0000Q4', ctx).center).toBe("R&D Q4");
  });

  it("supports page offsets", () => {
    expect(renderHeaderFooter("&P+10", ctx).center).toBe("12");
  });

  it("formats the date", () => {
    expect(renderHeaderFooter("&D", ctx).center).toBe(ctx.date.toLocaleDateString());
  });
});
