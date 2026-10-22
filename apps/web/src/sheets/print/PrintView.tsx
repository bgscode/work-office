// File → Print, like Excel's Print page: settings on the left (what to print,
// orientation, paper, margins, scaling, Page Setup), a live preview of one page
// on the right with page navigation, and Print. Printing uses the same HTML as
// the preview in a hidden frame, so the browser's "Save as PDF" printer gives
// the PDF export.

import { useEffect, useMemo, useRef, useState } from "react";
import type { SheetsDocument } from "@office/engine";
import { Icon } from "@office/ui";
import type { Rect } from "../geometry";
import { Select } from "../Select";
import { MARGIN_PRESETS, PAPER_SIZES, SCALING_OPTIONS, paperLabel, pageInches, readPageSetup, scalingChange, scalingId, updatePageSetup } from "./pageSetup";
import { PX_PER_INCH, buildPrintHtml, layoutJob } from "./printHtml";
import { buildPrintJob, type PrintWhat } from "./printJob";

/** Above this many pages, printing asks first (a big sheet can make thousands). */
const CONFIRM_PAGES = 200;

interface Props {
  doc: SheetsDocument;
  sheet: number;
  selection: Rect;
  title: string;
  revision: number;
  /** After a page setup change (undoable, so the workbook is saved). */
  onChanged: () => void;
  onPageSetup: () => void;
  onClose: () => void;
}

/** Print `html` through a hidden frame. */
export function printHtml(html: string) {
  const frame = document.createElement("iframe");
  frame.setAttribute("aria-hidden", "true");
  frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden";
  document.body.appendChild(frame);
  const done = () => setTimeout(() => frame.remove(), 1000);
  frame.onload = () => {
    const win = frame.contentWindow;
    if (!win) return done();
    win.addEventListener("afterprint", done);
    // Let images (charts) decode before the print dialog snapshots the page.
    setTimeout(() => {
      win.focus();
      win.print();
    }, 100);
  };
  frame.srcdoc = html;
}

export function PrintView({ doc, sheet, selection, title, revision, onChanged, onPageSetup, onClose }: Props) {
  const [what, setWhat] = useState<PrintWhat>("active");
  const [index, setIndex] = useState(0);
  const setup = useMemo(() => readPageSetup(doc, sheet), [doc, sheet, revision]);
  const job = useMemo(() => buildPrintJob(doc, sheet, what, selection, title), [doc, sheet, what, selection, title, revision]);
  const pages = useMemo(() => layoutJob(job), [job]);
  const page = Math.min(index, Math.max(0, pages.length - 1));
  const previewRef = useRef<HTMLDivElement>(null);
  const [box, setBox] = useState({ w: 600, h: 700 });

  useEffect(() => {
    const el = previewRef.current!;
    const ro = new ResizeObserver(() => setBox({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
      else if (e.key === "PageDown" || e.key === "ArrowRight") setIndex((i) => Math.min(pages.length - 1, i + 1));
      else if (e.key === "PageUp" || e.key === "ArrowLeft") setIndex((i) => Math.max(0, i - 1));
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [pages.length, onClose]);

  const pageSheet = pages[page] ? job.sheets[pages[page].sheet].setup : setup;
  const inches = pageInches(pageSheet);
  const pw = inches.width * PX_PER_INCH;
  const ph = inches.height * PX_PER_INCH;
  const scale = Math.min((box.w - 48) / pw, (box.h - 48) / ph, 1.5);
  const previewHtml = useMemo(() => (pages.length ? buildPrintHtml(job, pages, { include: [page], preview: true }) : ""), [job, pages, page]);

  function change(changes: Record<string, string>) {
    if (!updatePageSetup(doc, sheet, changes)) onChanged();
  }

  function print() {
    if (!pages.length) return;
    if (pages.length > CONFIRM_PAGES && !confirm(`This will print ${pages.length} pages. Continue?`)) return;
    printHtml(buildPrintHtml(job, pages));
  }

  const scaling = scalingId(setup);
  return (
    <div className="print-view" role="dialog" aria-label="Print">
      <aside className="print-settings">
        <button className="print-back" aria-label="Back" onClick={onClose}>
          <Icon name="chevronLeft" size={16} /> Print
        </button>
        <div className="print-actions">
          <button className="print-button" onClick={print} disabled={!pages.length}>
            <Icon name="print" size={32} />
            <span>Print</span>
          </button>
          <button className="print-button secondary" onClick={print} disabled={!pages.length} title="Choose “Save as PDF” as the printer in the dialog that opens.">
            <Icon name="download" size={32} />
            <span>Export PDF</span>
          </button>
        </div>
        <p className="print-hint">To save a PDF, pick “Save as PDF” as the destination in the print dialog.</p>
        <h3>Settings</h3>
        <Select aria-label="What to print" value={what} onChange={(e) => (setWhat(e.target.value as PrintWhat), setIndex(0))}>
          <option value="active">Print Active Sheets</option>
          <option value="workbook">Print Entire Workbook</option>
          <option value="selection">Print Selection</option>
        </Select>
        <Select aria-label="Orientation" value={setup.orientation} onChange={(e) => change({ orientation: e.target.value })}>
          <option value="portrait">Portrait Orientation</option>
          <option value="landscape">Landscape Orientation</option>
        </Select>
        <Select aria-label="Paper size" value={String(setup.paper)} onChange={(e) => change({ paper: e.target.value })}>
          {PAPER_SIZES.map((p) => (
            <option key={p.code} value={String(p.code)}>
              {p.name} ({paperLabel(p)})
            </option>
          ))}
        </Select>
        <Select aria-label="Margins" value={setup.margins.preset} onChange={(e) => change({ margins: e.target.value })}>
          {MARGIN_PRESETS.map((m) => (
            <option key={m.id} value={m.id}>
              {m.label} Margins
            </option>
          ))}
          {setup.margins.preset === "custom" && <option value="custom">Custom Margins</option>}
        </Select>
        <Select aria-label="Scaling" value={scaling} onChange={(e) => e.target.value !== "custom" && change(scalingChange(e.target.value as "none"))}>
          {SCALING_OPTIONS.map((o) => (
            <option key={o.id} value={o.id}>
              {o.label}
            </option>
          ))}
          {scaling === "custom" && <option value="custom">Custom Scaling</option>}
        </Select>
        <button className="print-link" onClick={onPageSetup}>
          Page Setup
        </button>
      </aside>
      <section className="print-preview">
        <div className="print-page-area" ref={previewRef}>
          {pages.length ? (
            <div className="print-page" style={{ width: pw * scale, height: ph * scale }}>
              <iframe
                title={`Page ${page + 1}`}
                srcDoc={previewHtml}
                style={{ width: pw, height: ph, transform: `scale(${scale})` }}
                tabIndex={-1}
              />
            </div>
          ) : (
            <p className="print-empty">We didn't find anything to print.</p>
          )}
        </div>
        <div className="print-nav">
          <button aria-label="Previous page" disabled={page <= 0} onClick={() => setIndex(page - 1)}>
            <Icon name="chevronLeft" size={12} />
          </button>
          <span>
            {pages.length ? page + 1 : 0} of {pages.length}
          </span>
          <button aria-label="Next page" disabled={page >= pages.length - 1} onClick={() => setIndex(page + 1)}>
            <Icon name="chevronRight" size={12} />
          </button>
        </div>
      </section>
    </div>
  );
}
