// Excel-style status bar: mode indicator (Ready / Enter / Edit / Point) on the
// left; Average, Count and Sum of the selection and the zoom slider on the right.

import type { SheetsDocument } from "@office/engine";
import { Icon, Tooltip } from "@office/ui";
import { MAX_ZOOM, MIN_ZOOM, sliderToZoom, stepZoom, zoomToSlider, type StatusMode } from "./chrome";
import type { Rect } from "./geometry";

/** Preset zoom levels in the percentage dropdown. */
const ZOOM_LEVELS = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2];

interface Props {
  doc: SheetsDocument;
  sheet: number;
  /** Selection limited to cells with data, or null. */
  rect: Rect | null;
  mode: StatusMode;
  /** Shown after the mode, e.g. "4 of 10 records found" while a filter is on. */
  note?: string;
  zoom: number;
  onZoom: (zoom: number) => void;
}

export function StatusBar({ doc, sheet, rect, mode, note, zoom, onZoom }: Props) {
  return (
    <div className="status-bar" role="status" aria-label="Status bar">
      <span className="status-mode" aria-label={`Mode: ${mode}`}>
        {mode}
      </span>
      {note && <span className="status-note">{note}</span>}
      <span className="status-spacer" />
      <SelectionStats doc={doc} sheet={sheet} rect={rect} />
      <ZoomControl zoom={zoom} onZoom={onZoom} />
    </div>
  );
}

function SelectionStats({ doc, sheet, rect }: { doc: SheetsDocument; sheet: number; rect: Rect | null }) {
  if (!rect || (rect.top === rect.bottom && rect.left === rect.right)) return <span className="status-stats" />;
  const [count, numbers, sum] = doc.rangeStats(sheet, rect.top, rect.left, rect.bottom - rect.top + 1, rect.right - rect.left + 1);
  const fmt = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 10 });
  return (
    <span className="status-stats">
      {numbers > 0 && <span>Average: {fmt(sum / numbers)}</span>}
      {count > 0 && <span>Count: {count}</span>}
      {numbers > 0 && <span>Sum: {fmt(sum)}</span>}
    </span>
  );
}

function ZoomControl({ zoom, onZoom }: { zoom: number; onZoom: (z: number) => void }) {
  const keep = (e: React.MouseEvent) => e.preventDefault();
  return (
    <span className="zoom-control">
      <Tooltip label="Zoom Out">
        <button className="zoom-step" aria-label="Zoom out" disabled={zoom <= MIN_ZOOM} onMouseDown={keep} onClick={() => onZoom(stepZoom(zoom, -1))}>
          <Icon name="minus" size={12} />
        </button>
      </Tooltip>
      <input
        className="zoom-slider"
        type="range"
        min={0}
        max={100}
        step={1}
        aria-label="Zoom slider"
        aria-valuetext={`${Math.round(zoom * 100)}%`}
        value={zoomToSlider(zoom)}
        onChange={(e) => onZoom(sliderToZoom(Number(e.target.value)))}
      />
      <Tooltip label="Zoom In">
        <button className="zoom-step" aria-label="Zoom in" disabled={zoom >= MAX_ZOOM} onMouseDown={keep} onClick={() => onZoom(stepZoom(zoom, 1))}>
          <Icon name="plus" size={12} />
        </button>
      </Tooltip>
      <select className="zoom" aria-label="Zoom" title="Zoom level (Ctrl + mouse wheel)" value={zoom} onChange={(e) => onZoom(Number(e.target.value))}>
        {[...new Set([...ZOOM_LEVELS, zoom])]
          .sort((a, b) => a - b)
          .map((z) => (
            <option key={z} value={z}>
              {Math.round(z * 100)}%
            </option>
          ))}
      </select>
    </span>
  );
}
