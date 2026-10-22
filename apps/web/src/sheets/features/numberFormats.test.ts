import { describe, expect, it } from "vitest";
import { accountingCode, categoryOf, currencyCode, decimalsOf, numberCode, percentCode, scientificCode } from "./numberFormats";

describe("Format Cells codes", () => {
  it("builds number codes with negative styles", () => {
    expect(numberCode(2, true, "minus")).toBe("#,##0.00");
    expect(numberCode(0, false, "red")).toBe("0;[Red]0");
    expect(numberCode(1, true, "parens")).toBe("#,##0.0_);(#,##0.0)");
  });

  it("builds currency, accounting, percent and scientific codes", () => {
    expect(currencyCode(2, "inr", "minus")).toBe("[$₹-4009]#,##0.00");
    expect(currencyCode(0, "usd", "redParens")).toBe('"$"#,##0_);[Red]("$"#,##0)');
    expect(accountingCode(2, "inr")).toBe('_([$₹-4009]* #,##0.00_);_([$₹-4009]* (#,##0.00);_([$₹-4009]* "-"??_);_(@_)');
    expect(percentCode(1)).toBe("0.0%");
    expect(scientificCode(2)).toBe("0.00E+00");
  });

  it("finds the category a code belongs to", () => {
    expect(categoryOf("General", undefined)).toBe("general");
    expect(categoryOf("dd-mmm-yyyy", "custom")).toBe("date");
    expect(categoryOf("0.00%", "percent")).toBe("percentage");
    expect(categoryOf('0" kg"', "custom")).toBe("custom");
    expect(decimalsOf('#,##0.000;"x"')).toBe(3);
  });
});
