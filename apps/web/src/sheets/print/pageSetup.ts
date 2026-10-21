// Page setup as the UI sees it: the JSON from `doc.pageSetup(sheet)` plus the
// option lists the Page Layout tab, Print page and Page Setup dialog offer.
// The engine owns the settings (crates/sheets-engine/src/page_setup.rs);
// changes go through `doc.setPageSetup`, which is one undo step.

import type { SheetsDocument } from "@office/engine";

export type Orientation = "portrait" | "landscape";
export type MarginPreset = "normal" | "wide" | "narrow" | "custom";

export interface Margins {
  preset: MarginPreset;
  /** Inches, as Excel stores them. */
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

/** Matches `PageSetup::to_json` in the engine. */
export interface PageSetup {
  orientation: Orientation;
  /** `.xlsx` paper size code (9 = A4, 1 = Letter, 5 = Legal...). */
  paper: number;
  /** Percent; only used when `fit` is null. */
  scale: number;
  /** Fit to [pages wide, pages tall]; 0 = automatic. */
  fit: [number, number] | null;
  margins: Margins;
  centerH: boolean;
  centerV: boolean;
  gridlines: boolean;
  headings: boolean;
  /** "A1:D20", "A:D" or "1:20". */
  printArea: string | null;
  /** "1:2" (one-based). */
  titleRows: string | null;
  /** "A:B". */
  titleCols: string | null;
  /** Excel header/footer codes, e.g. "&CPage &P of &N". */
  header: string;
  footer: string;
}

export function readPageSetup(doc: SheetsDocument, sheet: number): PageSetup {
  return JSON.parse(doc.pageSetup(sheet)) as PageSetup;
}

/**
 * Apply settings (keys as in `PageSetup::set`: orientation, paper, scale, fit,
 * margins, centerH, centerV, gridlines, headings, printArea, titleRows,
 * titleCols, header, footer) as one undo step. Returns an error message, or
 * null on success.
 */
export function updatePageSetup(doc: SheetsDocument, sheet: number, changes: Record<string, string>): string | null {
  const keys = Object.keys(changes);
  try {
    doc.setPageSetup(sheet, keys, keys.map((k) => changes[k]));
    return null;
  } catch (err) {
    return err instanceof Error ? err.message : String(err);
  }
}

export interface PaperSize {
  code: number;
  name: string;
  /** Portrait size in inches. */
  width: number;
  height: number;
}

const MM = 1 / 25.4;

/** The sizes Excel lists first, in its order. */
export const PAPER_SIZES: PaperSize[] = [
  { code: 1, name: "Letter", width: 8.5, height: 11 },
  { code: 3, name: "Tabloid", width: 11, height: 17 },
  { code: 5, name: "Legal", width: 8.5, height: 14 },
  { code: 7, name: "Executive", width: 7.25, height: 10.5 },
  { code: 8, name: "A3", width: 297 * MM, height: 420 * MM },
  { code: 9, name: "A4", width: 210 * MM, height: 297 * MM },
  { code: 11, name: "A5", width: 148 * MM, height: 210 * MM },
];

export function paperSize(code: number): PaperSize {
  return PAPER_SIZES.find((p) => p.code === code) ?? { code, name: `Paper ${code}`, width: 8.5, height: 11 };
}

/** "21 cm x 29.7 cm" like Excel's size gallery. */
export function paperLabel(p: PaperSize): string {
  const cm = (inches: number) => String(Math.round(inches * 2.54 * 100) / 100);
  return `${cm(p.width)} cm x ${cm(p.height)} cm`;
}

/** Page size in inches for the orientation. */
export function pageInches(setup: PageSetup): { width: number; height: number } {
  const p = paperSize(setup.paper);
  return setup.orientation === "landscape" ? { width: p.height, height: p.width } : { width: p.width, height: p.height };
}

export const MARGIN_PRESETS: { id: Exclude<MarginPreset, "custom">; label: string; m: Omit<Margins, "preset"> }[] = [
  { id: "normal", label: "Normal", m: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } },
  { id: "wide", label: "Wide", m: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } },
  { id: "narrow", label: "Narrow", m: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } },
];

/** Value for the `margins` key with custom inches. */
export function customMargins(m: Omit<Margins, "preset">): string {
  return [m.left, m.right, m.top, m.bottom, m.header, m.footer].join(",");
}

export type ScalingId = "none" | "sheet" | "columns" | "rows" | "custom";

export const SCALING_OPTIONS: { id: Exclude<ScalingId, "custom">; label: string; description: string }[] = [
  { id: "none", label: "No Scaling", description: "Print sheets at their actual size" },
  { id: "sheet", label: "Fit Sheet on One Page", description: "Shrink the printout so that it fits on one page" },
  { id: "columns", label: "Fit All Columns on One Page", description: "Shrink the printout so that it is one page wide" },
  { id: "rows", label: "Fit All Rows on One Page", description: "Shrink the printout so that it is one page high" },
];

export function scalingId(setup: PageSetup): ScalingId {
  if (setup.fit) {
    const [w, h] = setup.fit;
    if (w === 1 && h === 1) return "sheet";
    if (w === 1 && h === 0) return "columns";
    if (w === 0 && h === 1) return "rows";
    return "custom";
  }
  return setup.scale === 100 ? "none" : "custom";
}

export function scalingChange(id: Exclude<ScalingId, "custom">): Record<string, string> {
  switch (id) {
    case "none":
      return { scale: "100" };
    case "sheet":
      return { fit: "1,1" };
    case "columns":
      return { fit: "1,0" };
    case "rows":
      return { fit: "0,1" };
  }
}

/** Header/footer choices from Excel's dropdowns. `&A` sheet name, `&F` file name. */
export const HEADER_FOOTER_PRESETS: { label: string; code: string }[] = [
  { label: "(none)", code: "" },
  { label: "Page 1", code: "&CPage &P" },
  { label: "Page 1 of ?", code: "&CPage &P of &N" },
  { label: "Sheet1", code: "&C&A" },
  { label: "Book1", code: "&C&F" },
  { label: "Sheet1, Page 1", code: "&C&A, Page &P" },
  { label: "Book1, Page 1", code: "&C&F, Page &P" },
  { label: "Prepared by me, date, Page 1", code: "&LPrepared by me&C&D&RPage &P" },
];
