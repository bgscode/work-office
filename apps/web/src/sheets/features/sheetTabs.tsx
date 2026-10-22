// Sheet tab commands, like Excel's tab right-click menu and Home → Format:
// Move or Copy, Tab Color, Hide and Unhide. The engine keeps the order,
// colour and hidden state (crates/sheets-engine/src/ops.rs: MoveSheet, SetSheetTab).

import { useState } from "react";
import type { SheetsDocument } from "@office/engine";
import { Dialog } from "../Dialog";
import { FILL_COLORS } from "./FormatCellsDialog";

export interface SheetTab {
  /** `#RRGGBB`, or null for no colour. */
  color: string | null;
  hidden: boolean;
}

export function readTabs(doc: SheetsDocument): SheetTab[] {
  return JSON.parse(doc.sheetTabs()) as SheetTab[];
}

/** The next visible sheet from `index` in `direction`, or `index` when there is none. */
export function nextVisible(tabs: SheetTab[], index: number, direction: 1 | -1): number {
  for (let i = index + direction; i >= 0 && i < tabs.length; i += direction) {
    if (!tabs[i].hidden) return i;
  }
  return index;
}

/** Where a sheet goes when dropped before `before` (an index, or `length` for the end). */
export function moveTarget(from: number, before: number): number {
  return before > from ? before - 1 : before;
}

/** Black or white text, whichever reads better on `color`. */
export function textOn(color: string): string {
  const n = parseInt(color.slice(1), 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? "#000000" : "#FFFFFF";
}

interface TabColorProps {
  current: string | null;
  onPick: (color: string | null) => void;
  onClose: () => void;
}

/** Tab Color: Excel's palette, No Color, and More Colors. */
export function TabColorDialog({ current, onPick, onClose }: TabColorProps) {
  const pick = (c: string | null) => {
    onPick(c);
    onClose();
  };
  return (
    <Dialog title="Tab Color" onClose={onClose} width={300} footer={<button className="dialog-button" onClick={onClose}>Cancel</button>}>
      <div className="fill-swatches" role="listbox" aria-label="Tab colours">
        {FILL_COLORS.map((c) => (
          <button key={c} role="option" aria-selected={current === c} aria-label={c} className={current === c ? "active" : undefined} style={{ background: c }} onClick={() => pick(c)} />
        ))}
      </div>
      <div className="dialog-row">
        <button className="dialog-button" onClick={() => pick(null)}>
          No Color
        </button>
        <label className="dialog-field inline">
          <span>More Colors:</span>
          <input type="color" aria-label="Tab color" defaultValue={current ?? "#4472C4"} onChange={(e) => pick(e.target.value.toUpperCase())} />
        </label>
      </div>
    </Dialog>
  );
}

interface UnhideProps {
  names: string[];
  tabs: SheetTab[];
  onUnhide: (index: number) => void;
  onClose: () => void;
}

/** Unhide: pick one of the hidden sheets. */
export function UnhideDialog({ names, tabs, onUnhide, onClose }: UnhideProps) {
  const hidden = tabs.map((t, i) => (t.hidden ? i : -1)).filter((i) => i >= 0);
  const [chosen, setChosen] = useState(hidden[0] ?? -1);
  const ok = () => {
    if (chosen >= 0) onUnhide(chosen);
    onClose();
  };
  return (
    <Dialog
      title="Unhide"
      onClose={onClose}
      width={320}
      footer={
        <>
          <button className="dialog-button primary" disabled={chosen < 0} onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <p className="dialog-label">Unhide sheet:</p>
      <div className="dialog-list" role="listbox" aria-label="Hidden sheets">
        {hidden.map((i) => (
          <div key={i} role="option" aria-selected={chosen === i} className={chosen === i ? "current" : undefined} onMouseDown={() => setChosen(i)} onDoubleClick={() => (setChosen(i), onUnhide(i), onClose())}>
            {names[i]}
          </div>
        ))}
      </div>
    </Dialog>
  );
}

interface MoveCopyProps {
  names: string[];
  tabs: SheetTab[];
  index: number;
  /** `before`: an index, or names.length for "(move to end)". */
  onDone: (before: number, copy: boolean) => void;
  onClose: () => void;
}

/** Move or Copy: where to put the sheet, and whether to copy it. */
export function MoveCopyDialog({ names, tabs, index, onDone, onClose }: MoveCopyProps) {
  const [before, setBefore] = useState(Math.min(index + 1, names.length));
  const [copy, setCopy] = useState(false);
  const options = [...names.map((n, i) => ({ i, label: n })).filter(({ i }) => !tabs[i].hidden), { i: names.length, label: "(move to end)" }];
  const ok = () => {
    onDone(before, copy);
    onClose();
  };
  return (
    <Dialog
      title="Move or Copy"
      onClose={onClose}
      width={340}
      footer={
        <>
          <button className="dialog-button primary" onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <p className="dialog-label">Move selected sheet ({names[index]}) before sheet:</p>
      <div className="dialog-list" role="listbox" aria-label="Before sheet">
        {options.map(({ i, label }) => (
          <div key={i} role="option" aria-selected={before === i} className={before === i ? "current" : undefined} onMouseDown={() => setBefore(i)} onDoubleClick={() => (setBefore(i), onDone(i, copy), onClose())}>
            {label}
          </div>
        ))}
      </div>
      <label className="dialog-check">
        <input type="checkbox" checked={copy} onChange={(e) => setCopy(e.target.checked)} /> Create a copy
      </label>
    </Dialog>
  );
}
