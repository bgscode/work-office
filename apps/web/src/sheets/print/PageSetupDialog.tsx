// Excel's Page Setup dialog: Page (orientation, scaling, paper), Margins
// (inches, centring), Header/Footer (presets or custom codes) and Sheet
// (print area, print titles, gridlines, headings). OK applies everything as
// one undo step.

import { useState, type ReactNode } from "react";
import type { SheetsDocument } from "@office/engine";
import { Dialog } from "../Dialog";
import { Select } from "../Select";
import { HEADER_FOOTER_PRESETS, PAPER_SIZES, customMargins, paperLabel, readPageSetup, updatePageSetup } from "./pageSetup";

export type PageSetupTab = "page" | "margins" | "headerFooter" | "sheet";

interface Props {
  doc: SheetsDocument;
  sheet: number;
  initialTab?: PageSetupTab;
  onDone: () => void;
  onPrintPreview: () => void;
  onClose: () => void;
}

export function PageSetupDialog({ doc, sheet, initialTab = "page", onDone, onPrintPreview, onClose }: Props) {
  const start = readPageSetup(doc, sheet);
  const [tab, setTab] = useState<PageSetupTab>(initialTab);
  const [orientation, setOrientation] = useState(start.orientation);
  const [fitMode, setFitMode] = useState(start.fit ? "fit" : "scale");
  const [scale, setScale] = useState(String(start.scale));
  const [fitW, setFitW] = useState(String(start.fit?.[0] ?? 1));
  const [fitH, setFitH] = useState(start.fit?.[1] ? String(start.fit[1]) : "");
  const [paper, setPaper] = useState(String(start.paper));
  const m = start.margins;
  const [margins, setMargins] = useState({ top: m.top, bottom: m.bottom, left: m.left, right: m.right, header: m.header, footer: m.footer });
  const [centerH, setCenterH] = useState(start.centerH);
  const [centerV, setCenterV] = useState(start.centerV);
  const [header, setHeader] = useState(start.header);
  const [footer, setFooter] = useState(start.footer);
  const [printArea, setPrintArea] = useState(start.printArea ?? "");
  const [titleRows, setTitleRows] = useState(start.titleRows ?? "");
  const [titleCols, setTitleCols] = useState(start.titleCols ?? "");
  const [gridlines, setGridlines] = useState(start.gridlines);
  const [headings, setHeadings] = useState(start.headings);
  const [error, setError] = useState("");

  function apply(): boolean {
    const changes: Record<string, string> = {
      orientation,
      paper,
      centerH: centerH ? "1" : "0",
      centerV: centerV ? "1" : "0",
      header,
      footer,
      printArea: printArea.replace(/^=/, "").trim(),
      titleRows: titleRows.replace(/^=/, "").trim(),
      titleCols: titleCols.replace(/^=/, "").trim(),
      gridlines: gridlines ? "1" : "0",
      headings: headings ? "1" : "0",
      margins: customMargins(margins),
    };
    if (fitMode === "scale") changes.scale = scale;
    else changes.fit = `${Number(fitW) || 0},${Number(fitH) || 0}`;
    const err = updatePageSetup(doc, sheet, changes);
    if (err) {
      setError(err.replace(/^invalid value for (\w+).*$/, "Check the $1 setting."));
      return false;
    }
    onDone();
    return true;
  }

  const num = (label: string, value: number, set: (v: number) => void) => (
    <label className="dialog-field">
      <span>{label}</span>
      <input aria-label={label} inputMode="decimal" value={String(value)} onChange={(e) => set(Number(e.target.value) || 0)} />
    </label>
  );
  const check = (label: string, value: boolean, set: (v: boolean) => void) => (
    <label className="dialog-check">
      <input type="checkbox" checked={value} onChange={(e) => set(e.target.checked)} /> {label}
    </label>
  );
  const hfEditor = (label: string, value: string, set: (v: string) => void) => (
    <>
      <label className="dialog-field">
        <span>{label}:</span>
        <Select aria-label={label} value={HEADER_FOOTER_PRESETS.some((p) => p.code === value) ? value : "__custom"} onChange={(e) => e.target.value !== "__custom" && set(e.target.value)}>
          {HEADER_FOOTER_PRESETS.map((p) => (
            <option key={p.code} value={p.code}>
              {p.label}
            </option>
          ))}
          <option value="__custom">Custom…</option>
        </Select>
      </label>
      <label className="dialog-field">
        <span>Custom {label.toLowerCase()} (&amp;L left, &amp;C centre, &amp;R right, &amp;P page, &amp;N pages, &amp;D date, &amp;A sheet, &amp;F file):</span>
        <input aria-label={`Custom ${label.toLowerCase()}`} value={value} onChange={(e) => set(e.target.value)} />
      </label>
    </>
  );

  let body: ReactNode;
  if (tab === "page") {
    body = (
      <>
        <p className="dialog-label">Orientation</p>
        <div className="dialog-row radios">
          <label>
            <input type="radio" checked={orientation === "portrait"} onChange={() => setOrientation("portrait")} /> Portrait
          </label>
          <label>
            <input type="radio" checked={orientation === "landscape"} onChange={() => setOrientation("landscape")} /> Landscape
          </label>
        </div>
        <p className="dialog-label">Scaling</p>
        <div className="dialog-row">
          <label>
            <input type="radio" checked={fitMode === "scale"} onChange={() => setFitMode("scale")} /> Adjust to:
          </label>
          <input aria-label="Scale percent" inputMode="numeric" style={{ width: 60, flex: "none" }} value={scale} onChange={(e) => setScale(e.target.value)} onFocus={() => setFitMode("scale")} />
          <span>% normal size</span>
        </div>
        <div className="dialog-row">
          <label>
            <input type="radio" checked={fitMode === "fit"} onChange={() => setFitMode("fit")} /> Fit to:
          </label>
          <input aria-label="Pages wide" inputMode="numeric" style={{ width: 50, flex: "none" }} value={fitW} onChange={(e) => setFitW(e.target.value)} onFocus={() => setFitMode("fit")} />
          <span>page(s) wide by</span>
          <input aria-label="Pages tall" inputMode="numeric" placeholder="auto" style={{ width: 50, flex: "none" }} value={fitH} onChange={(e) => setFitH(e.target.value)} onFocus={() => setFitMode("fit")} />
          <span>tall</span>
        </div>
        <label className="dialog-field">
          <span>Paper size:</span>
          <Select aria-label="Paper size" value={paper} onChange={(e) => setPaper(e.target.value)}>
            {PAPER_SIZES.map((p) => (
              <option key={p.code} value={String(p.code)}>
                {p.name} ({paperLabel(p)})
              </option>
            ))}
          </Select>
        </label>
      </>
    );
  } else if (tab === "margins") {
    const set = (k: keyof typeof margins) => (v: number) => setMargins((x) => ({ ...x, [k]: v }));
    body = (
      <>
        <p className="dialog-label">Margins in inches</p>
        <div className="dialog-grid3">
          {num("Top", margins.top, set("top"))}
          {num("Header", margins.header, set("header"))}
          {num("Left", margins.left, set("left"))}
          {num("Right", margins.right, set("right"))}
          {num("Bottom", margins.bottom, set("bottom"))}
          {num("Footer", margins.footer, set("footer"))}
        </div>
        <p className="dialog-label">Center on page</p>
        {check("Horizontally", centerH, setCenterH)}
        {check("Vertically", centerV, setCenterV)}
      </>
    );
  } else if (tab === "headerFooter") {
    body = (
      <>
        {hfEditor("Header", header, setHeader)}
        {hfEditor("Footer", footer, setFooter)}
      </>
    );
  } else {
    body = (
      <>
        <label className="dialog-field">
          <span>Print area:</span>
          <input aria-label="Print area" placeholder="A1:H40" value={printArea} onChange={(e) => setPrintArea(e.target.value)} />
        </label>
        <p className="dialog-label">Print titles</p>
        <label className="dialog-field">
          <span>Rows to repeat at top:</span>
          <input aria-label="Rows to repeat at top" placeholder="$1:$1" value={titleRows} onChange={(e) => setTitleRows(e.target.value)} />
        </label>
        <label className="dialog-field">
          <span>Columns to repeat at left:</span>
          <input aria-label="Columns to repeat at left" placeholder="$A:$A" value={titleCols} onChange={(e) => setTitleCols(e.target.value)} />
        </label>
        <p className="dialog-label">Print</p>
        {check("Gridlines", gridlines, setGridlines)}
        {check("Row and column headings", headings, setHeadings)}
      </>
    );
  }

  const tabs: [PageSetupTab, string][] = [
    ["page", "Page"],
    ["margins", "Margins"],
    ["headerFooter", "Header/Footer"],
    ["sheet", "Sheet"],
  ];
  return (
    <Dialog
      title="Page Setup"
      onClose={onClose}
      width={480}
      footer={
        <>
          <button
            className="dialog-button push-left"
            onClick={() => {
              if (apply()) {
                onClose();
                onPrintPreview();
              }
            }}
          >
            Print Preview
          </button>
          <button className="dialog-button primary" onClick={() => apply() && onClose()}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <div className="dialog-tabs" role="tablist">
        {tabs.map(([id, label]) => (
          <button key={id} role="tab" aria-selected={tab === id} className={tab === id ? "dialog-tab active" : "dialog-tab"} onClick={() => setTab(id)}>
            {label}
          </button>
        ))}
      </div>
      <div className="dialog-tab-body">{body}</div>
      {error && <p className="dialog-error">{error}</p>}
    </Dialog>
  );
}
