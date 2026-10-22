// The Number tab of Format Cells: categories like Excel's and the format
// codes they produce. Pure, so the codes are unit-tested; the engine formats
// with any Excel code (crates/sheets-engine/src/numfmt.rs).

export type Category =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "date"
  | "time"
  | "percentage"
  | "fraction"
  | "scientific"
  | "text"
  | "custom";

export const CATEGORIES: { id: Category; label: string; description: string }[] = [
  { id: "general", label: "General", description: "General format cells have no specific number format." },
  { id: "number", label: "Number", description: "Number is used for general display of numbers. Currency and Accounting offer specialised formatting for monetary value." },
  { id: "currency", label: "Currency", description: "Currency formats are used for general monetary values." },
  { id: "accounting", label: "Accounting", description: "Accounting formats line up the currency symbols and decimal points in a column." },
  { id: "date", label: "Date", description: "Date formats display date and time serial numbers as date values." },
  { id: "time", label: "Time", description: "Time formats display date and time serial numbers as time values." },
  { id: "percentage", label: "Percentage", description: "Percentage formats multiply the cell value by 100 and display the result with a percent symbol." },
  { id: "fraction", label: "Fraction", description: "Fraction formats show the number as a fraction." },
  { id: "scientific", label: "Scientific", description: "Scientific formats show the number in exponential notation." },
  { id: "text", label: "Text", description: "Text format cells are treated as text even when a number is in the cell." },
  { id: "custom", label: "Custom", description: "Type the number format code, using one of the existing codes as a starting point." },
];

export const SYMBOLS: { id: string; label: string; code: string }[] = [
  { id: "inr", label: "₹ Indian Rupee", code: "[$₹-4009]" },
  { id: "usd", label: "$ US Dollar", code: '"$"' },
  { id: "eur", label: "€ Euro", code: '"€"' },
  { id: "gbp", label: "£ Pound", code: '"£"' },
  { id: "jpy", label: "¥ Yen", code: '"¥"' },
  { id: "none", label: "None", code: "" },
];

/** How negative numbers show (Number and Currency). */
export type NegativeStyle = "minus" | "red" | "parens" | "redParens";
export const NEGATIVE_STYLES: { id: NegativeStyle; label: string; color?: string }[] = [
  { id: "minus", label: "-1234.10" },
  { id: "red", label: "1234.10", color: "#FF0000" },
  { id: "parens", label: "(1234.10)" },
  { id: "redParens", label: "(1234.10)", color: "#FF0000" },
];

export const DATE_CODES = [
  "dd-mm-yyyy",
  "d-m-yy",
  "dd/mm/yyyy",
  "d-mmm-yy",
  "dd-mmm-yyyy",
  "d mmmm yyyy",
  "dddd, d mmmm yyyy",
  "mmm-yy",
  "mmmm yyyy",
  "d-mmm",
  "yyyy-mm-dd",
  "m/d/yyyy",
  "dd-mm-yyyy hh:mm",
];

export const TIME_CODES = ["hh:mm", "h:mm AM/PM", "hh:mm:ss", "h:mm:ss AM/PM", "mm:ss.0", "[h]:mm:ss", "dd-mm-yyyy h:mm AM/PM"];

export const FRACTION_CODES: { code: string; label: string }[] = [
  { code: "# ?/?", label: "Up to one digit (1/4)" },
  { code: "# ??/??", label: "Up to two digits (21/25)" },
  { code: "# ???/???", label: "Up to three digits (312/943)" },
  { code: "# ?/2", label: "As halves (1/2)" },
  { code: "# ?/4", label: "As quarters (2/4)" },
  { code: "# ?/8", label: "As eighths (4/8)" },
  { code: "# ??/16", label: "As sixteenths (8/16)" },
  { code: "# ?/10", label: "As tenths (3/10)" },
  { code: "# ??/100", label: "As hundredths (30/100)" },
];

/** Codes offered in the Custom list (Excel's built-ins and common business ones). */
export const CUSTOM_CODES = [
  "General",
  "0",
  "0.00",
  "#,##0",
  "#,##0.00",
  "#,##0;-#,##0",
  "#,##0;[Red]-#,##0",
  "#,##0.00;(#,##0.00)",
  "#,##0.00;[Red](#,##0.00)",
  '[$₹-4009]#,##,##0.00',
  '#,##0,"K"',
  '#,##0.0,,"M"',
  "0%",
  "0.00%",
  "0.00E+00",
  "# ?/?",
  "dd-mm-yyyy",
  "dd-mmm-yyyy",
  "mmm-yy",
  "h:mm AM/PM",
  "[h]:mm:ss",
  '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  '0" kg"',
  '"Rs. "#,##0',
  "@",
];

const decimalsPart = (d: number) => (d > 0 ? "." + "0".repeat(d) : "");

/** Number category: `#,##0.00`, with the negative style. */
export function numberCode(decimals: number, thousands: boolean, negative: NegativeStyle, prefix = ""): string {
  const body = prefix + (thousands ? "#,##0" : "0") + decimalsPart(decimals);
  switch (negative) {
    case "minus":
      return body;
    case "red":
      return `${body};[Red]${body}`;
    case "parens":
      return `${body}_);(${body})`;
    case "redParens":
      return `${body}_);[Red](${body})`;
  }
}

export function currencyCode(decimals: number, symbol: string, negative: NegativeStyle): string {
  const sym = SYMBOLS.find((s) => s.id === symbol)?.code ?? "";
  return numberCode(decimals, true, negative, sym);
}

/** Accounting: symbol at the left edge, a dash for zero, negatives in parentheses. */
export function accountingCode(decimals: number, symbol: string): string {
  const sym = SYMBOLS.find((s) => s.id === symbol)?.code ?? "";
  const n = "#,##0" + decimalsPart(decimals);
  const dash = decimals > 0 ? `"-"${"?".repeat(decimals)}` : '"-"';
  return `_(${sym}* ${n}_);_(${sym}* (${n});_(${sym}* ${dash}_);_(@_)`;
}

export function percentCode(decimals: number): string {
  return "0" + decimalsPart(decimals) + "%";
}

export function scientificCode(decimals: number): string {
  return "0" + decimalsPart(decimals) + "E+00";
}

/** The category and settings a code most likely came from, to open the dialog on it. */
export function categoryOf(code: string, fmt: string | undefined): Category {
  if (!code || code === "General") return "general";
  if (code === "@") return "text";
  if (DATE_CODES.includes(code) || fmt === "date" || fmt === "datetime") return "date";
  if (TIME_CODES.includes(code) || fmt === "time") return "time";
  if (FRACTION_CODES.some((f) => f.code === code)) return "fraction";
  if (/E\+0+$/.test(code)) return "scientific";
  if (/^0(\.0+)?%$/.test(code)) return "percentage";
  if (code.startsWith("_(")) return "accounting";
  if (fmt === "currency") return "currency";
  if (fmt === "number") return "number";
  return "custom";
}

/** Decimal places in the first section of a code. */
export function decimalsOf(code: string): number {
  const first = code.split(";")[0].replace(/"[^"]*"/g, "");
  const m = /\.(0+)/.exec(first);
  return m ? m[1].length : 0;
}
