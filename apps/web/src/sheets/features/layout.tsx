// Ribbon and menu commands for the sheet's layout: Merge & Center, Freeze
// Panes, row heights / column widths and Hide & Unhide. SheetsApp owns the
// state and passes callbacks in; this file only builds commands and dialogs.

import { useState } from "react";
import type { RibbonCommand, RibbonMenuItem } from "@office/ui";
import { Dialog } from "../Dialog";
import type { Rect } from "../geometry";
import type { MergeIndex } from "../merges";
import type { Freeze } from "../panes";

export type MergeKind = "center" | "across" | "cells";
export type Axis = "rows" | "cols";

export interface LayoutContext {
  sel: Rect;
  merges: MergeIndex;
  freeze: Freeze;
  /** Active cell, for Freeze Panes at the selection. */
  active: { row: number; col: number };
  merge: (kind: MergeKind) => void;
  unmerge: () => void;
  setFreeze: (rows: number, cols: number) => void;
  setHidden: (axis: Axis, hidden: boolean) => void;
  openSizeDialog: (axis: Axis) => void;
  autofit: (axis: Axis) => void;
}

/** Excel shows row heights in points and column widths in characters of the default font. */
export const pxToPoints = (px: number) => Math.round(px * 0.75 * 100) / 100;
export const pointsToPx = (pt: number) => Math.round(pt / 0.75);
export const pxToChars = (px: number) => Math.round(((px - 5) / 7) * 100) / 100;
export const charsToPx = (chars: number) => Math.round(chars * 7 + 5);

export function mergeCommand(ctx: LayoutContext): RibbonCommand {
  const merged = ctx.merges.within(ctx.sel).length > 0;
  return {
    id: "merge",
    label: "Merge & Center",
    title: "Combine the selected cells into one larger cell and center the contents.",
    icon: "merge",
    size: "small",
    group: "Alignment",
    active: merged,
    onRun: () => (merged ? ctx.unmerge() : ctx.merge("center")),
    menu: [
      { id: "merge-center", label: "Merge & Center", icon: "merge", onRun: () => ctx.merge("center") },
      { id: "merge-across", label: "Merge Across", onRun: () => ctx.merge("across") },
      { id: "merge-cells", label: "Merge Cells", onRun: () => ctx.merge("cells") },
      { id: "merge-undo", label: "Unmerge Cells", onRun: ctx.unmerge, disabled: !merged },
    ],
  };
}

export function freezeCommand(ctx: LayoutContext): RibbonCommand {
  const frozen = ctx.freeze.rows > 0 || ctx.freeze.cols > 0;
  const { row, col } = ctx.active;
  return {
    id: "freeze",
    label: "Freeze Panes",
    title: "Keep a portion of the sheet visible while the rest of the sheet scrolls.",
    icon: "freeze",
    size: "large",
    group: "Window",
    menu: [
      frozen
        ? { id: "unfreeze", label: "Unfreeze Panes", onRun: () => ctx.setFreeze(0, 0) }
        : { id: "freeze-here", label: "Freeze Panes", onRun: () => ctx.setFreeze(row, col), disabled: row === 0 && col === 0 },
      { id: "freeze-top", label: "Freeze Top Row", onRun: () => ctx.setFreeze(1, 0) },
      { id: "freeze-first", label: "Freeze First Column", onRun: () => ctx.setFreeze(0, 1) },
    ],
  };
}

/** Items for Home → Cells → Format (Excel's Cell Size and Visibility sections). */
export function formatMenuItems(ctx: LayoutContext): RibbonMenuItem[] {
  return [
    { id: "menu-row-height", label: "Row Height…", onRun: () => ctx.openSizeDialog("rows") },
    { id: "menu-autofit-rows", label: "AutoFit Row Height", onRun: () => ctx.autofit("rows") },
    { id: "menu-col-width", label: "Column Width…", onRun: () => ctx.openSizeDialog("cols"), separatorBefore: true },
    { id: "menu-autofit-cols", label: "AutoFit Column Width", onRun: () => ctx.autofit("cols") },
    { id: "menu-hide-rows", label: "Hide Rows", onRun: () => ctx.setHidden("rows", true), separatorBefore: true },
    { id: "menu-hide-cols", label: "Hide Columns", onRun: () => ctx.setHidden("cols", true) },
    { id: "menu-unhide-rows", label: "Unhide Rows", onRun: () => ctx.setHidden("rows", false) },
    { id: "menu-unhide-cols", label: "Unhide Columns", onRun: () => ctx.setHidden("cols", false) },
  ];
}

interface SizeDialogProps {
  axis: Axis;
  /** Current size in pixels (of the first selected line). */
  px: number;
  onApply: (px: number) => void;
  onClose: () => void;
}

/** Excel's Row Height (points) / Column Width (characters) dialog. */
export function SizeDialog({ axis, px, onApply, onClose }: SizeDialogProps) {
  const isRows = axis === "rows";
  const [value, setValue] = useState(String(isRows ? pxToPoints(px) : pxToChars(px)));
  const max = isRows ? 409 : 255;
  const n = Number(value);
  const valid = value.trim() !== "" && n >= 0 && n <= max;
  function ok() {
    if (!valid) return;
    onApply(isRows ? pointsToPx(n) : charsToPx(n));
    onClose();
  }
  return (
    <Dialog
      title={isRows ? "Row Height" : "Column Width"}
      onClose={onClose}
      width={280}
      footer={
        <>
          <button className="dialog-button primary" disabled={!valid} onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <label className="dialog-field inline">
        <span>{isRows ? "Row height:" : "Column width:"}</span>
        <input
          autoFocus
          inputMode="decimal"
          aria-label={isRows ? "Row height in points" : "Column width in characters"}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && ok()}
          style={{ width: 80 }}
        />
      </label>
      {!valid && <p className="dialog-error">Enter a number between 0 and {max}.</p>}
    </Dialog>
  );
}
