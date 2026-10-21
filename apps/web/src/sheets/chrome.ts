// Pure helpers behind the Excel-style chrome (status bar, zoom slider, Insert
// Function). No React or engine imports, so they are unit-tested in chrome.test.ts.

import { HEADER_HEIGHT, HEADER_WIDTH } from "./geometry";
import type { EditState } from "./types";

/** Excel's status bar mode indicator. */
export type StatusMode = "Ready" | "Enter" | "Edit" | "Point";

export function statusMode(editing: EditState | null, pointMode: boolean): StatusMode {
  if (!editing) return "Ready";
  if (pointMode) return "Point";
  return editing.mode === "enter" ? "Enter" : "Edit";
}

export const MIN_ZOOM = 0.5;
export const MAX_ZOOM = 2;

/**
 * The zoom slider puts 100% in the middle, like Excel: the left half covers
 * 50–100% and the right half 100–200%. Slider positions run 0–100.
 */
export function zoomToSlider(zoom: number): number {
  const z = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom));
  return z <= 1 ? ((z - MIN_ZOOM) / (1 - MIN_ZOOM)) * 50 : 50 + ((z - 1) / (MAX_ZOOM - 1)) * 50;
}

export function sliderToZoom(position: number): number {
  const p = Math.min(100, Math.max(0, position));
  const z = p <= 50 ? MIN_ZOOM + (p / 50) * (1 - MIN_ZOOM) : 1 + ((p - 50) / 50) * (MAX_ZOOM - 1);
  // Snap to whole percents, and to 100% near the middle.
  const pct = Math.round(z * 100);
  return Math.abs(pct - 100) <= 2 ? 1 : pct / 100;
}

/** The − / + buttons step to the next multiple of 10%, like Excel. */
export function stepZoom(zoom: number, direction: 1 | -1): number {
  const pct = Math.round(zoom * 100);
  const next = direction > 0 ? Math.floor(pct / 10) * 10 + 10 : Math.ceil(pct / 10) * 10 - 10;
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, next / 100));
}

/**
 * View → Zoom to Selection: the largest zoom (whole percent, 50–200%) at which
 * a `width`×`height` block of cells plus the headers fits a `viewW`×`viewH` grid.
 */
export function zoomToFit(width: number, height: number, viewW: number, viewH: number): number {
  const z = Math.min(viewW / (HEADER_WIDTH + width), viewH / (HEADER_HEIGHT + height));
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Math.floor(z * 100) / 100));
}

/**
 * Insert `NAME(` for the Insert Function dialog. Inside a formula it goes at
 * the caret; otherwise the entry is replaced by `=NAME(`.
 */
export function insertFunctionText(draft: string | null, caret: number, name: string): { draft: string; caret: number } {
  if (draft === null || !draft.startsWith("=")) {
    const text = `=${name}(`;
    return { draft: text, caret: text.length };
  }
  const at = Math.min(Math.max(caret, 1), draft.length);
  const next = draft.slice(0, at) + name + "(" + draft.slice(at);
  return { draft: next, caret: at + name.length + 1 };
}

/** Function names matching an Insert Function search, best matches first. */
export function searchFunctions<T extends { name: string; summary: string }>(all: T[], query: string): T[] {
  const q = query.trim().toUpperCase();
  if (!q) return all;
  const starts = all.filter((f) => f.name.startsWith(q));
  const contains = all.filter((f) => !f.name.startsWith(q) && f.name.includes(q));
  const words = q.toLowerCase().split(/\s+/);
  const described = all.filter(
    (f) => !f.name.includes(q) && words.every((w) => f.summary.toLowerCase().includes(w)),
  );
  return [...starts, ...contains, ...described];
}

/** "Ctrl+" on Windows/Linux, "⌘" on Mac, for shortcut labels. */
export const MOD_KEY = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";
