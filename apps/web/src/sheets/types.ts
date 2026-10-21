import type { Selection } from "./geometry";

/**
 * An edit in progress. `source` says which input has focus. In "enter" mode
 * (started by typing over a cell) arrow keys commit and move, like Excel;
 * in "edit" mode (F2, double-click, formula bar) they move the caret.
 */
export interface EditState {
  draft: string;
  source: "cell" | "bar";
  mode: "enter" | "edit";
  /** Caret position in `draft`. */
  caret: number;
  /** The reference most recently inserted by clicking/arrowing (point mode). */
  point?: { start: number; end: number; sel: Selection };
}

export const FORMULA_BAR_ID = "formula-bar-input";
