// The Page Layout tab: Margins, Orientation, Size, Print Area, Print Titles,
// Scale to Fit and Sheet Options (print gridlines/headings), like Excel's.
// Settings live in the engine (page_setup.rs) and change as one undo step.

import type { RibbonCommand } from "@office/ui";
import { Select } from "../Select";
import type { PageSetupTab } from "../print/PageSetupDialog";
import { MARGIN_PRESETS, PAPER_SIZES, paperLabel, type PageSetup } from "../print/pageSetup";

export interface PageLayoutContext {
  setup: PageSetup;
  change: (changes: Record<string, string>) => void;
  /** The selection as "A1:D20", for Set Print Area. */
  selectionRef: string;
  openPageSetup: (tab: PageSetupTab) => void;
}

const FIT_CHOICES = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"];

export function pageLayoutCommands(ctx: PageLayoutContext): RibbonCommand[] {
  const s = ctx.setup;
  const fitW = s.fit ? String(s.fit[0]) : "0";
  const fitH = s.fit ? String(s.fit[1]) : "0";
  const pages = (n: string) => (n === "0" ? "Automatic" : n === "1" ? "1 page" : `${n} pages`);
  const setFit = (w: string, h: string) => ctx.change(w === "0" && h === "0" ? { scale: "100" } : { fit: `${w},${h}` });
  return [
    {
      id: "pl-margins",
      label: "Margins",
      title: "Set the margins for the entire document or the current section.",
      icon: "margins",
      size: "large",
      group: "Page Setup",
      menu: [
        ...MARGIN_PRESETS.map((m) => ({
          id: `margins-${m.id}`,
          label: `${m.label}  (Top ${m.m.top}" Bottom ${m.m.bottom}" Left ${m.m.left}" Right ${m.m.right}")`,
          active: s.margins.preset === m.id,
          onRun: () => ctx.change({ margins: m.id }),
        })),
        { id: "margins-custom", label: "Custom Margins…", active: s.margins.preset === "custom", onRun: () => ctx.openPageSetup("margins"), separatorBefore: true },
      ],
    },
    {
      id: "pl-orientation",
      label: "Orientation",
      title: "Switch the pages between portrait and landscape layouts.",
      icon: "orientation",
      size: "large",
      group: "Page Setup",
      menu: [
        { id: "orient-portrait", label: "Portrait", active: s.orientation === "portrait", onRun: () => ctx.change({ orientation: "portrait" }) },
        { id: "orient-landscape", label: "Landscape", active: s.orientation === "landscape", onRun: () => ctx.change({ orientation: "landscape" }) },
      ],
    },
    {
      id: "pl-size",
      label: "Size",
      title: "Choose a paper size for the printout.",
      icon: "paperSize",
      size: "large",
      group: "Page Setup",
      menu: PAPER_SIZES.map((p) => ({ id: `size-${p.code}`, label: `${p.name}  ${paperLabel(p)}`, active: s.paper === p.code, onRun: () => ctx.change({ paper: String(p.code) }) })),
    },
    {
      id: "pl-print-area",
      label: "Print Area",
      title: "Mark a specific area of the sheet for printing.",
      icon: "printArea",
      size: "large",
      group: "Page Setup",
      menu: [
        { id: "print-area-set", label: "Set Print Area", onRun: () => ctx.change({ printArea: ctx.selectionRef }) },
        { id: "print-area-clear", label: "Clear Print Area", disabled: !s.printArea, onRun: () => ctx.change({ printArea: "" }) },
      ],
    },
    {
      id: "pl-titles",
      label: "Print Titles",
      title: "Specify rows and columns to repeat on each printed page.",
      icon: "printTitles",
      size: "large",
      group: "Page Setup",
      onRun: () => ctx.openPageSetup("sheet"),
    },
    {
      id: "pl-setup",
      label: "Page Setup…",
      title: "Show all page setup options.",
      icon: "pageSetup",
      size: "small",
      group: "Page Setup",
      onRun: () => ctx.openPageSetup("page"),
    },
    {
      id: "pl-width",
      label: "Width",
      title: "Shrink the width of the printout to fit a number of pages.",
      group: "Scale to Fit",
      render: () => (
        <label className="ribbon-field">
          <span>Width:</span>
          <Select aria-label="Fit width" value={fitW} onChange={(e) => setFit(e.target.value, fitH)}>
            {FIT_CHOICES.map((n) => (
              <option key={n} value={n}>
                {pages(n)}
              </option>
            ))}
          </Select>
        </label>
      ),
    },
    {
      id: "pl-height",
      label: "Height",
      newRow: true,
      title: "Shrink the height of the printout to fit a number of pages.",
      group: "Scale to Fit",
      render: () => (
        <label className="ribbon-field">
          <span>Height:</span>
          <Select aria-label="Fit height" value={fitH} onChange={(e) => setFit(fitW, e.target.value)}>
            {FIT_CHOICES.map((n) => (
              <option key={n} value={n}>
                {pages(n)}
              </option>
            ))}
          </Select>
        </label>
      ),
    },
    {
      id: "pl-scale",
      label: "Scale",
      newRow: true,
      title: "Stretch or shrink the printout to a percentage of its actual size. Width and Height must be Automatic.",
      group: "Scale to Fit",
      render: () => (
        <label className="ribbon-field">
          <span>Scale:</span>
          <Select aria-label="Scale" value={s.fit ? "" : String(s.scale)} disabled={!!s.fit} onChange={(e) => ctx.change({ scale: e.target.value })}>
            {s.fit && <option value="">Fit</option>}
            {[...new Set([25, 50, 75, 85, 90, 100, 110, 125, 150, 200, 400, s.scale])]
              .sort((a, b) => a - b)
              .map((p) => (
                <option key={p} value={String(p)}>
                  {p}%
                </option>
              ))}
          </Select>
        </label>
      ),
    },
    {
      id: "pl-gridlines",
      label: "Print Gridlines",
      title: "Print the lines between rows and columns.",
      icon: s.gridlines ? "check" : "gridlines",
      size: "small",
      group: "Sheet Options",
      active: s.gridlines,
      onRun: () => ctx.change({ gridlines: s.gridlines ? "0" : "1" }),
    },
    {
      id: "pl-headings",
      label: "Print Headings",
      title: "Print row numbers and column letters.",
      icon: s.headings ? "check" : "headings",
      size: "small",
      group: "Sheet Options",
      active: s.headings,
      onRun: () => ctx.change({ headings: s.headings ? "0" : "1" }),
    },
  ];
}
