import { describe, expect, it } from "vitest";
import type { PageSetup } from "./pageSetup";
import { buildPrintHtml, colLetters, colSpan, layoutJob, rowSpan, type PrintJob, type PrintStyle, type SheetSource } from "./printHtml";

const setup = (over: Partial<PageSetup> = {}): PageSetup => ({
  orientation: "portrait",
  paper: 9,
  scale: 100,
  fit: null,
  margins: { preset: "normal", left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  centerH: false,
  centerV: false,
  gridlines: false,
  headings: false,
  printArea: null,
  titleRows: null,
  titleCols: null,
  header: "",
  footer: "",
  ...over,
});

/** A sheet whose cell (r, c) holds `cells[r][c]`; numbers are numbers. */
function source(cells: (string | number)[][], s: PageSetup, styles: Record<string, PrintStyle> = {}): SheetSource {
  return {
    name: "Sheet1",
    setup: s,
    range: { top: 0, left: 0, bottom: cells.length - 1, right: Math.max(...cells.map((r) => r.length)) - 1 },
    colWidth: () => 100,
    rowHeight: () => 22,
    block: (r) => {
      const texts: string[] = [];
      const kinds: number[] = [];
      const st: PrintStyle[] = [];
      for (let row = r.top; row <= r.bottom; row++) {
        for (let col = r.left; col <= r.right; col++) {
          const v = cells[row]?.[col] ?? "";
          texts.push(String(v));
          kinds.push(v === "" ? 0 : typeof v === "number" ? 1 : 2);
          st.push(styles[`${row},${col}`] ?? {});
        }
      }
      return { texts, kinds, styles: st };
    },
  };
}

const job = (sheets: SheetSource[]): PrintJob => ({
  title: "Book1",
  sheets,
  date: new Date(2025, 0, 1),
  measure: (text) => text.length * 7,
});

describe("span parsing", () => {
  it("reads print title spans", () => {
    expect(rowSpan("$1:$2")).toEqual([0, 1]);
    expect(rowSpan("3")).toEqual([2, 2]);
    expect(colSpan("B:A")).toEqual([0, 1]);
    expect(colSpan("AA")).toEqual([26, 26]);
    expect(rowSpan(null)).toBeNull();
    expect(colLetters(27)).toBe("AB");
  });
});

describe("buildPrintHtml", () => {
  it("prints cells with their formatting and alignment", () => {
    const j = job([source([["Item", 12.5], ["<b>&", ""]], setup(), { "0,0": { b: 1, color: "#C00000", fill: "#FFF2CC", align: "center" } })]);
    const pages = layoutJob(j);
    expect(pages.length).toBe(1);
    const html = buildPrintHtml(j, pages);
    expect(html).toContain('<td style="font-weight:bold;color:#C00000;background:#FFF2CC;text-align:center">Item</td>');
    expect(html).toContain('<td style="text-align:right">12.5</td>');
    expect(html).toContain("&lt;b&gt;&amp;");
    expect(html).toContain("@page s0{size:8.26771653543307in 11.69291338582677in;margin:0}");
  });

  it("shows ### for numbers that don't fit and spills long text", () => {
    const long = "A long heading that needs room";
    const html = buildPrintHtml(job([source([[123456789012345], [long, "", "x"]], setup())]), layoutJob(job([source([[123456789012345], [long, "", "x"]], setup())])));
    expect(html).toContain(">############</td>");
    expect(html).toContain(`<td colspan="2">${long}</td>`);
  });

  it("repeats title rows and numbers pages in the footer", () => {
    const rows = Array.from({ length: 80 }, (_, i) => [i === 0 ? "Header" : `r${i}`]);
    const j = job([source(rows, setup({ titleRows: "1:1", footer: "&CPage &P of &N" }))]);
    const pages = layoutJob(j);
    expect(pages.length).toBe(2);
    expect(pages[1].row.titles).toEqual([0, 0]);
    const html = buildPrintHtml(j, pages);
    expect(html.match(/>Header</g)?.length).toBe(2);
    expect(html).toContain('<div class="c">Page 2 of 2</div>');
  });

  it("prints gridlines and headings when asked", () => {
    const j = job([source([["a", "b"]], setup({ gridlines: true, headings: true }))]);
    const html = buildPrintHtml(j, layoutJob(j));
    expect(html).toContain('<table class="grid"');
    expect(html).toContain("<th>A</th><th>B</th>");
    expect(html).toContain("<th>1</th>");
  });

  it("uses landscape page size and scaling", () => {
    const wide = [Array.from({ length: 30 }, (_, i) => `c${i}`)];
    const j = job([source(wide, setup({ orientation: "landscape", fit: [1, 0] }))]);
    const pages = layoutJob(j);
    expect(pages.length).toBe(1);
    expect(pages[0].scale).toBeLessThan(0.4);
    expect(buildPrintHtml(j, pages)).toContain("size:11.69291338582677in 8.26771653543307in");
  });

  it("skips sheets with nothing to print", () => {
    const empty = { ...source([["x"]], setup()), range: null };
    expect(layoutJob(job([empty]))).toEqual([]);
  });

  it("includes only the requested pages", () => {
    const rows = Array.from({ length: 80 }, (_, i) => [`r${i}`]);
    const j = job([source(rows, setup())]);
    const html = buildPrintHtml(j, layoutJob(j), { include: [1], preview: true });
    expect(html).toContain('data-page="2"');
    expect(html).not.toContain('data-page="1"');
    expect(html).not.toContain("box-shadow");
  });
});
