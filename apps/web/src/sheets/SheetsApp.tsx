// Owns all spreadsheet UI state and turns user actions into engine calls.
// The engine (`SheetsDocument`, Rust via WebAssembly) owns the data; React
// state here is only selection, the edit in progress, menus, and a revision
// counter that tells the grid to repaint.

import { useEffect, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent } from "react";
import { SheetsDocument, functionList, parseStyle, type FunctionInfo } from "@office/engine";
import { Ribbon, type RibbonTab } from "@office/ui";
import { download, pickFile } from "../files";
import { clearDraft, loadDraft, saveDraft } from "./autosave";
import { parseTsv, toTsv } from "./clipboard";
import { ContextMenu, type MenuItem } from "./ContextMenu";
import { FindBar, type FindOptions } from "./FindBar";
import { FormulaAssist } from "./FormulaAssist";
import { FormulaBar } from "./FormulaBar";
import { canInsertRef, closeParens, enclosingFunction, findRefs, wordAtCaret } from "./formulaText";
import {
  MAX_COLS,
  MAX_ROWS,
  clampToExtent,
  moveSelection,
  rangeName,
  selectionRect,
  single,
  type Cell,
  type Rect,
  type Selection,
} from "./geometry";
import { Grid, fillArea, type FillDirection, type GridShortcut } from "./Grid";
import { fontFor } from "./paint";
import { NameBox } from "./NameBox";
import { SheetTabs } from "./SheetTabs";
import type { EditState } from "./types";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
/** Listed first in autocomplete when they match, since most people want these. */
const POPULAR = ["SUM", "AVERAGE", "COUNT", "IF", "VLOOKUP", "XLOOKUP", "SUMIF", "COUNTIF", "MAX", "MIN", "ROUND", "IFERROR", "TODAY", "CONCAT", "INDEX", "MATCH"];
const popularity = (name: string) => {
  const i = POPULAR.indexOf(name);
  return i === -1 ? POPULAR.length : i;
};

const ZOOM_KEY = "office.sheets.zoom";
const ZOOM_LEVELS = [0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2];

function savedZoom(): number {
  try {
    const z = Number(localStorage.getItem(ZOOM_KEY));
    return z >= 0.5 && z <= 2 ? z : 1;
  } catch {
    return 1;
  }
}

/** Formatting more cells than this at once is limited to the area holding data. */
const MAX_FORMAT_CELLS = 200_000;

function openInitialDocument(): { doc: SheetsDocument; name: string } {
  const draft = loadDraft();
  if (draft) {
    try {
      return { doc: SheetsDocument.fromXlsx(draft.xlsx), name: draft.name };
    } catch {
      clearDraft();
    }
  }
  return { doc: new SheetsDocument(), name: "Untitled" };
}

export function SheetsApp() {
  const [initial] = useState(openInitialDocument);
  const [doc, setDoc] = useState(initial.doc);
  const [fileName, setFileName] = useState(initial.name);
  const [sheet, setSheet] = useState(0);
  const [selection, setSelection] = useState<Selection>(single({ row: 0, col: 0 }));
  const [editing, setEditing] = useState<EditState | null>(null);
  const [revision, setRevision] = useState(0);
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null);
  const [findMode, setFindMode] = useState<"find" | "replace" | null>(null);
  const [zoom, setZoomState] = useState(savedZoom);
  const [reveal, setReveal] = useState<{ row: number; col: number; seq: number } | null>(null);
  const [renamingTab, setRenamingTab] = useState<number | null>(null);
  const [assistIndex, setAssistIndex] = useState(0);
  const [assistDismissedAt, setAssistDismissedAt] = useState<number | null>(null);
  const cellInputRef = useRef<HTMLTextAreaElement>(null);
  const barInputRef = useRef<HTMLTextAreaElement>(null);
  const nameBoxRef = useRef<HTMLInputElement>(null);
  /** Set by Ctrl+Shift+V just before the browser's paste event. */
  const pasteValuesNext = useRef(false);
  const pendingCaret = useRef<number | null>(null);
  const lastCopiedText = useRef<string | null>(null);
  const changed = () => setRevision((r) => r + 1);

  const { row, col } = selection.anchor;
  const sel = selectionRect(selection);
  const activeStyle = parseStyle(doc.styleBlock(sheet, row, col, 1, 1)[0]);

  /** Select and scroll into view. */
  function goTo(s: Selection) {
    setSelection(s);
    setReveal((r) => ({ row: s.focus.row, col: s.focus.col, seq: (r?.seq ?? 0) + 1 }));
  }

  function setZoom(z: number) {
    setZoomState(z);
    try {
      localStorage.setItem(ZOOM_KEY, String(z));
    } catch {
      // ignore
    }
  }

  // ---- autosave: a second after the last change

  useEffect(() => {
    if (revision === 0) return;
    const timer = setTimeout(() => saveDraft(fileName, doc.toXlsx()), 1000);
    return () => clearTimeout(timer);
  }, [doc, revision, fileName]);

  // ---- editing

  function startEdit(source: EditState["source"], initialText?: string) {
    const draft = initialText ?? doc.input(sheet, row, col);
    setEditing({ draft, source, mode: initialText === undefined ? "edit" : "enter", caret: draft.length });
    setAssistDismissedAt(null);
    setAssistIndex(0);
  }

  function changeDraft(draft: string, caret: number) {
    setEditing((e) => e && { ...e, draft, caret, point: undefined });
    setAssistIndex(0);
  }

  function commit(dRow: number, dCol: number) {
    if (!editing) return;
    const value = closeParens(editing.draft);
    if (value !== doc.input(sheet, row, col)) {
      doc.setInputs(sheet, Uint32Array.of(row), Uint32Array.of(col), [value]);
      changed();
    }
    setEditing(null);
    if (dRow || dCol) setSelection(moveSelection(selection, dRow, dCol, false));
  }

  /** Ctrl+Enter: put the entry in every selected cell; formulas shift like a fill. */
  function commitToSelection() {
    if (!editing) return;
    let r: Rect | null = sel;
    if ((r.bottom - r.top + 1) * (r.right - r.left + 1) > MAX_FORMAT_CELLS) r = dataRect();
    const value = closeParens(editing.draft);
    if (r) {
      const isFormula = value.startsWith("=");
      setRect(r, (rr, cc) => (isFormula ? SheetsDocument.shiftFormula(value, rr - row, cc - col) : value));
    }
    setEditing(null);
  }

  // Put the caret where a programmatic draft change (autocomplete, point mode) left it.
  useLayoutEffect(() => {
    const caret = pendingCaret.current;
    if (caret === null || !editing) return;
    pendingCaret.current = null;
    const el = editing.source === "cell" ? cellInputRef.current : barInputRef.current;
    el?.focus({ preventScroll: true });
    el?.setSelectionRange(caret, caret);
  });

  function replaceDraft(draft: string, caret: number, point?: EditState["point"]) {
    if (!editing) return;
    setEditing({ ...editing, draft, caret, point });
    pendingCaret.current = caret;
  }

  // ---- formula assistance: references, autocomplete, point mode

  const isFormula = !!editing && editing.draft.startsWith("=");
  const highlights = useMemo(() => (isFormula ? findRefs(editing!.draft) : []), [isFormula, editing?.draft]);
  const pointMode =
    isFormula && (canInsertRef(editing!.draft, editing!.caret) || editing!.point?.end === editing!.caret);

  const word = isFormula ? wordAtCaret(editing!.draft, editing!.caret) : null;
  const suggestions: FunctionInfo[] =
    word && assistDismissedAt !== word.start && editing!.draft[editing!.caret] !== "("
      ? functionList()
          .filter((f) => f.name.startsWith(word.text.toUpperCase()))
          .sort((a, b) => popularity(a.name) - popularity(b.name) || a.name.localeCompare(b.name))
          .slice(0, 8)
      : [];
  const hintName = isFormula ? enclosingFunction(editing!.draft, editing!.caret) : null;
  const hint = hintName ? (functionList().find((f) => f.name === hintName) ?? null) : null;
  const activeInput = editing?.source === "bar" ? barInputRef.current : editing ? cellInputRef.current : null;

  function acceptSuggestion(f: FunctionInfo) {
    if (!editing || !word) return;
    const draft = editing.draft.slice(0, word.start) + f.name + "(" + editing.draft.slice(editing.caret);
    replaceDraft(draft, word.start + f.name.length + 1);
  }

  /** Keys handled before the editor's own: line breaks, Ctrl+Enter, autocomplete. */
  function onEditorKey(e: KeyboardEvent<HTMLTextAreaElement>): boolean {
    if (!editing) return false;
    if (e.key === "Enter" && e.altKey) {
      // Alt+Enter: line break inside the cell.
      e.preventDefault();
      const { draft, caret } = editing;
      replaceDraft(draft.slice(0, caret) + "\n" + draft.slice(caret), caret + 1);
      return true;
    }
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      commitToSelection();
      return true;
    }
    if (suggestions.length === 0) return false;
    const index = Math.min(assistIndex, suggestions.length - 1);
    switch (e.key) {
      case "ArrowDown":
        e.preventDefault();
        setAssistIndex((index + 1) % suggestions.length);
        return true;
      case "ArrowUp":
        e.preventDefault();
        setAssistIndex((index - 1 + suggestions.length) % suggestions.length);
        return true;
      case "Tab":
      case "Enter":
        e.preventDefault();
        acceptSuggestion(suggestions[index]);
        return true;
      case "Escape":
        e.preventDefault();
        setAssistDismissedAt(word!.start);
        return true;
    }
    return false;
  }

  /** Insert (or, while dragging, update) a reference to `s` at the caret. */
  function point(s: Selection, phase: "start" | "drag") {
    if (!editing) return;
    const text = rangeName(s);
    const replacing = editing.point && (phase === "drag" || editing.caret === editing.point.end);
    const start = replacing ? editing.point!.start : editing.caret;
    const end = replacing ? editing.point!.end : editing.caret;
    const draft = editing.draft.slice(0, start) + text + editing.draft.slice(end);
    replaceDraft(draft, start + text.length, { start, end: start + text.length, sel: s });
  }

  function pointMove(dRow: number, dCol: number, extend: boolean) {
    if (!editing) return;
    const base = editing.point?.sel ?? single(selection.anchor);
    point(moveSelection(base, dRow, dCol, extend), "start");
  }

  // ---- range helpers

  /** The selection limited to cells that hold data, or null if it holds none. */
  function dataRect(): Rect | null {
    const [rows, cols] = doc.extent(sheet);
    return clampToExtent(sel, rows, cols);
  }

  function setRect(r: Rect, values: (row: number, col: number) => string) {
    const rows: number[] = [];
    const cols: number[] = [];
    const inputs: string[] = [];
    for (let rr = r.top; rr <= r.bottom; rr++) {
      for (let cc = r.left; cc <= r.right; cc++) {
        rows.push(rr);
        cols.push(cc);
        inputs.push(values(rr, cc));
      }
    }
    doc.setInputs(sheet, Uint32Array.from(rows), Uint32Array.from(cols), inputs);
    changed();
  }

  function clear() {
    const r = dataRect();
    if (r) setRect(r, () => "");
  }

  // ---- clipboard

  function copy(): string {
    const r = dataRect() ?? { top: row, left: col, bottom: row, right: col };
    const width = r.right - r.left + 1;
    const flat = doc.copy(sheet, r.top, r.left, r.bottom, r.right);
    const rows: string[][] = [];
    for (let i = 0; i < flat.length; i += width) rows.push(flat.slice(i, i + width));
    const text = toTsv(rows);
    lastCopiedText.current = text;
    return text;
  }

  function cut(): string {
    const text = copy();
    clear();
    return text;
  }

  function paste(text: string) {
    const valuesOnly = pasteValuesNext.current;
    pasteValuesNext.current = false;
    // Our own copy: paste formulas (shifted) and formatting, not just values.
    if (lastCopiedText.current !== null && text === lastCopiedText.current) {
      pasteSpecial(valuesOnly ? "values" : "all", false);
      return;
    }
    const data = parseTsv(text);
    if (data.length === 0) return;
    const width = Math.max(...data.map((r) => r.length));
    const r: Rect = {
      top: row,
      left: col,
      bottom: Math.min(row + data.length - 1, MAX_ROWS - 1),
      right: Math.min(col + width - 1, MAX_COLS - 1),
    };
    setRect(r, (rr, cc) => data[rr - row]?.[cc - col] ?? "");
    setSelection({ anchor: { row: r.top, col: r.left }, focus: { row: r.bottom, col: r.right } });
  }

  /** Paste the last copy made in Office: everything, values, formulas or formats; optionally transposed. */
  function pasteSpecial(mode: "all" | "values" | "formulas" | "formats", transpose: boolean) {
    if (!doc.hasCopied()) {
      alert("Copy some cells in this workbook first (Ctrl+C), then use Paste Special.");
      return;
    }
    const size = doc.pasteCopied(sheet, row, col, mode, transpose);
    if (size.length !== 2) return;
    changed();
    setSelection({ anchor: { row, col }, focus: { row: row + size[0] - 1, col: col + size[1] - 1 } });
  }

  async function pasteFromMenu() {
    try {
      paste(await navigator.clipboard.readText());
    } catch {
      alert("Your browser blocked clipboard access here. Use Ctrl+V (⌘V on Mac) to paste.");
    }
  }

  // ---- formatting

  function format(key: string, value: string) {
    let r: Rect | null = sel;
    if ((r.bottom - r.top + 1) * (r.right - r.left + 1) > MAX_FORMAT_CELLS) r = dataRect();
    if (!r) return;
    doc.formatRange(sheet, r.top, r.left, r.bottom, r.right, key, value);
    changed();
  }

  const toggle = (key: "bold" | "italic" | "underline" | "strike", on: boolean | undefined) => format(key, on ? "0" : "1");

  // ---- fill handle

  function fill(direction: FillDirection, count: number, increment: boolean) {
    doc.fill(sheet, sel.top, sel.left, sel.bottom, sel.right, direction, count, increment);
    changed();
    const area = fillArea(sel, direction, count);
    setSelection({
      anchor: selection.anchor,
      focus: {
        row: direction === "down" ? area.bottom : direction === "up" ? area.top : selection.focus.row,
        col: direction === "right" ? area.right : direction === "left" ? area.left : selection.focus.col,
      },
    });
  }

  /** Double-click on the fill handle: fill down as far as the neighbouring column has data. */
  function fillToEnd() {
    const [rows, cols] = doc.extent(sheet);
    const neighbour = sel.left > 0 ? sel.left - 1 : sel.right + 1;
    if (neighbour >= cols || sel.bottom + 1 >= rows) return;
    const kinds = doc.kindBlock(sheet, sel.bottom + 1, neighbour, rows - sel.bottom - 1, 1);
    let count = 0;
    while (count < kinds.length && kinds[count] !== 0) count++;
    if (count > 0) fill("down", count, false);
  }

  function fillWithin(direction: "down" | "right") {
    // Ctrl+D / Ctrl+R: copy the first row/column of the selection across it.
    if (direction === "down" && sel.bottom > sel.top) {
      doc.fill(sheet, sel.top, sel.left, sel.top, sel.right, "down", sel.bottom - sel.top, false);
      changed();
    } else if (direction === "right" && sel.right > sel.left) {
      doc.fill(sheet, sel.top, sel.left, sel.bottom, sel.left, "right", sel.right - sel.left, false);
      changed();
    }
  }

  // ---- columns

  function resizeColumn(c: number, px: number) {
    // Resizing one of several selected whole columns resizes them all.
    const wholeCols = sel.top === 0 && sel.bottom === MAX_ROWS - 1 && c >= sel.left && c <= sel.right;
    doc.setColumnWidth(sheet, wholeCols ? sel.left : c, wholeCols ? sel.right : c, px);
    changed();
  }

  function autofit(c: number) {
    const [rows] = doc.extent(sheet);
    const n = Math.min(rows, 5000);
    if (n === 0) return;
    const texts = doc.displayBlock(sheet, 0, c, n, 1);
    const styles = doc.styleBlock(sheet, 0, c, n, 1);
    const ctx = document.createElement("canvas").getContext("2d")!;
    let widest = 0;
    texts.forEach((t, i) => {
      if (!t) return;
      ctx.font = fontFor(parseStyle(styles[i]));
      widest = Math.max(widest, ctx.measureText(t).width);
    });
    if (widest > 0) resizeColumn(c, Math.ceil(widest + 12));
  }

  // ---- rows/columns structure, sort

  const lineSpan = (axis: "rows" | "cols") =>
    axis === "rows" ? { at: sel.top, count: sel.bottom - sel.top + 1 } : { at: sel.left, count: sel.right - sel.left + 1 };

  function insertLines(axis: "rows" | "cols", after: boolean) {
    const { at, count } = lineSpan(axis);
    const capped = Math.min(count, axis === "rows" ? 10_000 : 1000);
    doc.insertLines(sheet, axis, after ? at + count : at, capped);
    changed();
  }

  function deleteLines(axis: "rows" | "cols") {
    const { at, count } = lineSpan(axis);
    doc.deleteLines(sheet, axis, at, count);
    changed();
    setSelection(single({ row: axis === "rows" ? sel.top : row, col: axis === "cols" ? sel.left : col }));
  }

  function sort(ascending: boolean) {
    const [rows, cols] = doc.extent(sheet);
    if (rows === 0) return;
    // One cell selected: sort the whole table, keeping row 1 as the header.
    const single = sel.top === sel.bottom && sel.left === sel.right;
    const r = single ? { top: 1, left: 0, bottom: rows - 1, right: cols - 1 } : dataRect();
    if (!r || r.bottom <= r.top) return;
    doc.sortRange(sheet, r.top, r.left, r.bottom, r.right, col, ascending);
    changed();
  }

  // ---- find

  function find(query: string, forward: boolean, opts: FindOptions): boolean {
    const hit = doc.find(sheet, query, row, col, forward, opts.matchCase, opts.wholeCell);
    if (hit.length !== 2) return false;
    goTo(single({ row: hit[0], col: hit[1] }));
    return true;
  }

  function replace(query: string, replacement: string, opts: FindOptions) {
    if (doc.replaceIn(sheet, row, col, query, replacement, opts.matchCase, opts.wholeCell)) changed();
    find(query, true, opts);
  }

  function replaceAll(query: string, replacement: string, opts: FindOptions): number {
    const n = doc.replaceAll(sheet, query, replacement, opts.matchCase, opts.wholeCell);
    if (n > 0) changed();
    return n;
  }

  /** Ctrl+arrow (Shift extends the selection). */
  function jump(dRow: number, dCol: number, extend: boolean) {
    const from = extend ? selection.focus : selection.anchor;
    const [r, c] = doc.jump(sheet, from.row, from.col, dRow, dCol);
    const to: Cell = { row: r, col: c };
    goTo(extend ? { anchor: selection.anchor, focus: to } : single(to));
  }

  function shortcut(s: GridShortcut) {
    const actions: Record<GridShortcut, () => void> = {
      bold: () => toggle("bold", !!activeStyle.b),
      italic: () => toggle("italic", !!activeStyle.i),
      underline: () => toggle("underline", !!activeStyle.u),
      find: () => setFindMode("find"),
      replace: () => setFindMode("replace"),
      goto: () => nameBoxRef.current?.focus(),
      pasteValuesNext: () => {
        pasteValuesNext.current = true;
      },
      fillDown: () => fillWithin("down"),
      fillRight: () => fillWithin("right"),
    };
    actions[s]();
  }

  function undo() {
    if (doc.undo()) changed();
  }

  function redo() {
    if (doc.redo()) changed();
  }

  // ---- files

  function replaceDocument(next: SheetsDocument, name: string) {
    doc.free();
    setDoc(next);
    setFileName(name);
    setSheet(0);
    setSelection(single({ row: 0, col: 0 }));
    setEditing(null);
    changed();
  }

  async function open() {
    const file = await pickFile(".xlsx,.csv");
    if (!file) return;
    const base = file.name.replace(/\.[^.]+$/, "");
    try {
      if (file.name.toLowerCase().endsWith(".csv")) {
        replaceDocument(SheetsDocument.fromCsv(await file.text()), base);
      } else {
        replaceDocument(SheetsDocument.fromXlsx(new Uint8Array(await file.arrayBuffer())), base);
      }
    } catch (err) {
      alert(`Could not open ${file.name}: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  function newDocument() {
    if (!confirm("Start a new, empty workbook? Download the current one first if you want to keep it.")) return;
    clearDraft();
    replaceDocument(new SheetsDocument(), "Untitled");
  }

  // ---- menus

  function cellMenu(x: number, y: number) {
    const mod = navigator.platform.startsWith("Mac") ? "⌘" : "Ctrl+";
    const items: MenuItem[] = [
      { label: "Cut", shortcut: `${mod}X`, onRun: () => navigator.clipboard?.writeText(cut()) },
      { label: "Copy", shortcut: `${mod}C`, onRun: () => navigator.clipboard?.writeText(copy()) },
      { label: "Paste", shortcut: `${mod}V`, onRun: pasteFromMenu },
      { label: "Paste values only", shortcut: `${mod}⇧V`, onRun: () => pasteSpecial("values", false), disabled: !doc.hasCopied() },
      { label: "Paste formatting only", onRun: () => pasteSpecial("formats", false), disabled: !doc.hasCopied() },
      { label: "Paste formulas only", onRun: () => pasteSpecial("formulas", false), disabled: !doc.hasCopied() },
      { label: "Paste transposed", onRun: () => pasteSpecial("all", true), disabled: !doc.hasCopied() },
      { label: "Insert row above", onRun: () => insertLines("rows", false), separatorBefore: true },
      { label: "Insert row below", onRun: () => insertLines("rows", true) },
      { label: "Insert column left", onRun: () => insertLines("cols", false) },
      { label: "Insert column right", onRun: () => insertLines("cols", true) },
      { label: "Delete rows", onRun: () => deleteLines("rows"), separatorBefore: true },
      { label: "Delete columns", onRun: () => deleteLines("cols") },
      { label: "Sort A → Z", onRun: () => sort(true), separatorBefore: true },
      { label: "Sort Z → A", onRun: () => sort(false) },
      { label: "Clear contents", shortcut: "Del", onRun: clear, separatorBefore: true },
      { label: "Clear formatting", onRun: () => format("clear", "") },
    ];
    setMenu({ x, y, items });
  }

  function tabMenu(index: number, x: number, y: number) {
    setMenu({
      x,
      y,
      items: [
        { label: "Rename", onRun: () => setRenamingTab(index) },
        {
          label: "Delete sheet",
          disabled: doc.sheetNames().length <= 1,
          onRun: () => {
            if (!confirm(`Delete "${doc.sheetNames()[index]}"? You can undo this.`)) return;
            if (doc.deleteSheet(index)) {
              setSheet(Math.max(0, Math.min(sheet, doc.sheetNames().length - 1)));
              changed();
            }
          },
        },
      ],
    });
  }

  // ---- ribbon

  const formatKind = activeStyle.fmt ?? "general";
  const tabs: RibbonTab[] = [
    {
      id: "file",
      label: "File",
      commands: [
        { id: "new", label: "New", onRun: newDocument },
        { id: "open", label: "Open…", onRun: open },
        { id: "save-xlsx", label: "Download .xlsx", onRun: () => download(new Uint8Array(doc.toXlsx()), `${fileName}.xlsx`, XLSX_MIME) },
        { id: "save-csv", label: "Download .csv", onRun: () => download(doc.toCsv(sheet), `${fileName}.csv`, "text/csv") },
      ],
    },
    {
      id: "home",
      label: "Home",
      commands: [
        { id: "undo", label: "↶", title: "Undo (Ctrl+Z)", onRun: undo },
        { id: "redo", label: "↷", title: "Redo (Ctrl+Y)", onRun: redo },
        {
          id: "paste-special",
          label: "Paste special",
          separatorBefore: true,
          render: () => (
            <select
              className="ribbon-select"
              aria-label="Paste special"
              value=""
              onChange={(e) => {
                const [mode, t] = e.target.value.split(":");
                pasteSpecial(mode as "all" | "values" | "formulas" | "formats", t === "t");
              }}
            >
              <option value="" disabled>
                Paste special…
              </option>
              <option value="values">Values only</option>
              <option value="formats">Formatting only</option>
              <option value="formulas">Formulas only</option>
              <option value="all:t">Transpose</option>
            </select>
          ),
        },
        { id: "bold", label: "B", title: "Bold (Ctrl+B)", active: !!activeStyle.b, onRun: () => toggle("bold", !!activeStyle.b), separatorBefore: true },
        { id: "italic", label: "I", title: "Italic (Ctrl+I)", active: !!activeStyle.i, onRun: () => toggle("italic", !!activeStyle.i) },
        { id: "underline", label: "U", title: "Underline (Ctrl+U)", active: !!activeStyle.u, onRun: () => toggle("underline", !!activeStyle.u) },
        { id: "strike", label: "S", title: "Strikethrough", active: !!activeStyle.s, onRun: () => toggle("strike", !!activeStyle.s) },
        {
          id: "color",
          label: "Text colour",
          render: () => <ColorPicker label="A" title="Text colour" value={activeStyle.color ?? "#1f1f1f"} onPick={(c) => format("color", c)} />,
        },
        {
          id: "fill",
          label: "Fill colour",
          render: () => <ColorPicker label="▧" title="Fill colour" value={activeStyle.fill ?? "#ffffff"} onPick={(c) => format("fill", c)} />,
        },
        { id: "align-left", label: "⇤", title: "Align left", active: activeStyle.align === "left", onRun: () => format("align", "left"), separatorBefore: true },
        { id: "align-center", label: "↔", title: "Centre", active: activeStyle.align === "center", onRun: () => format("align", "center") },
        { id: "align-right", label: "⇥", title: "Align right", active: activeStyle.align === "right", onRun: () => format("align", "right") },
        {
          id: "format",
          label: "Number format",
          separatorBefore: true,
          render: () => (
            <select className="ribbon-select" aria-label="Number format" value={formatKind} onChange={(e) => format("format", e.target.value)}>
              <option value="general">General</option>
              <option value="number">Number (1,234.50)</option>
              <option value="currency">Rupee (₹1,234.50)</option>
              <option value="dollar">Dollar ($1,234.50)</option>
              <option value="percent">Percent (12%)</option>
              <option value="date">Date (15-08-2025)</option>
              <option value="time">Time (14:30)</option>
              <option value="datetime">Date &amp; time</option>
              <option value="text">Text</option>
            </select>
          ),
        },
        { id: "dec-less", label: ".0←", title: "Fewer decimal places", onRun: () => format("decimals", "-1") },
        { id: "dec-more", label: ".00→", title: "More decimal places", onRun: () => format("decimals", "1") },
        { id: "clear-format", label: "Clear format", onRun: () => format("clear", ""), separatorBefore: true },
      ],
    },
    {
      id: "insert",
      label: "Insert",
      commands: [
        { id: "row-above", label: "Row above", onRun: () => insertLines("rows", false) },
        { id: "row-below", label: "Row below", onRun: () => insertLines("rows", true) },
        { id: "col-left", label: "Column left", onRun: () => insertLines("cols", false) },
        { id: "col-right", label: "Column right", onRun: () => insertLines("cols", true) },
        { id: "del-rows", label: "Delete rows", onRun: () => deleteLines("rows"), separatorBefore: true },
        { id: "del-cols", label: "Delete columns", onRun: () => deleteLines("cols") },
        { id: "sheet", label: "New sheet", onRun: addSheet, separatorBefore: true },
      ],
    },
    {
      id: "data",
      label: "Data",
      commands: [
        { id: "sort-asc", label: "Sort A → Z", title: "Sort by the active column", onRun: () => sort(true) },
        { id: "sort-desc", label: "Sort Z → A", title: "Sort by the active column", onRun: () => sort(false) },
        { id: "find", label: "Find", title: "Find (Ctrl+F)", onRun: () => setFindMode("find"), separatorBefore: true },
        { id: "replace", label: "Replace", title: "Find and replace (Ctrl+H)", onRun: () => setFindMode("replace") },
        { id: "goto", label: "Go to", title: "Go to a cell (Ctrl+G)", onRun: () => nameBoxRef.current?.focus() },
      ],
    },
  ];

  function addSheet() {
    setSheet(doc.addSheet());
    setSelection(single({ row: 0, col: 0 }));
    changed();
  }

  return (
    <div className="sheets-app">
      <Ribbon tabs={tabs} initialTab="home" />
      <FormulaBar
        inputRef={barInputRef}
        nameBox={
          <NameBox
            ref={nameBoxRef}
            label={rangeName(selection)}
            onGo={goTo}
            onDone={() => cellInputRef.current?.focus({ preventScroll: true })}
          />
        }
        input={doc.input(sheet, row, col)}
        editing={editing}
        onStartEdit={() => startEdit("bar")}
        onDraftChange={changeDraft}
        onCaret={(caret) => setEditing((e) => e && { ...e, caret })}
        onEditorKey={onEditorKey}
        onCommit={commit}
        onCancel={() => setEditing(null)}
      />
      <div className="grid-area">
        <Grid
          doc={doc}
          sheet={sheet}
          revision={revision}
          selection={selection}
          editing={editing}
          inputRef={cellInputRef}
          zoom={zoom}
          onZoom={setZoom}
          reveal={reveal}
          onJump={jump}
          pointMode={pointMode}
          highlights={highlights}
          onSelect={setSelection}
          onStartEdit={(text) => startEdit("cell", text)}
          onDraftChange={changeDraft}
          onCaret={(caret) => setEditing((e) => e && { ...e, caret })}
          onEditorKey={onEditorKey}
          onCommit={commit}
          onCancel={() => setEditing(null)}
          onPoint={point}
          onPointMove={pointMove}
          onClear={clear}
          onUndo={undo}
          onRedo={redo}
          onCopy={copy}
          onCut={cut}
          onPaste={paste}
          onFill={fill}
          onFillToEnd={fillToEnd}
          onResizeColumn={resizeColumn}
          onAutofit={autofit}
          onContextMenu={cellMenu}
          onShortcut={shortcut}
        />
        {findMode && (
          <FindBar
            withReplace={findMode === "replace"}
            onFind={find}
            onReplace={replace}
            onReplaceAll={replaceAll}
            onClose={() => {
              setFindMode(null);
              cellInputRef.current?.focus({ preventScroll: true });
            }}
          />
        )}
      </div>
      <footer className="sheets-footer">
        <SheetTabs
          names={doc.sheetNames()}
          active={sheet}
          onSelect={(i) => {
            if (editing) commit(0, 0);
            setSheet(i);
            setSelection(single({ row: 0, col: 0 }));
          }}
          onAdd={addSheet}
          onRename={(i, name) => {
            if (name === doc.sheetNames()[i]) return true;
            const ok = doc.renameSheet(i, name);
            if (ok) changed();
            return ok;
          }}
          onContextMenu={tabMenu}
          renaming={renamingTab}
          setRenaming={setRenamingTab}
        />
        <StatusBar doc={doc} sheet={sheet} rect={dataRect()} />
        <select className="zoom" aria-label="Zoom" title="Zoom (Ctrl + mouse wheel)" value={zoom} onChange={(e) => setZoom(Number(e.target.value))}>
          {[...new Set([...ZOOM_LEVELS, zoom])]
            .sort((a, b) => a - b)
            .map((z) => (
              <option key={z} value={z}>
                {Math.round(z * 100)}%
              </option>
            ))}
        </select>
      </footer>
      <FormulaAssist
        anchor={activeInput?.getBoundingClientRect() ?? null}
        suggestions={suggestions}
        selected={Math.min(assistIndex, Math.max(0, suggestions.length - 1))}
        hint={hint}
        onPick={acceptSuggestion}
        onHover={setAssistIndex}
      />
      {menu && <ContextMenu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} />}
    </div>
  );
}

/** A toolbar colour swatch. Applies the colour once the picker closes. */
function ColorPicker({ label, title, value, onPick }: { label: string; title: string; value: string; onPick: (c: string) => void }) {
  const ref = useRef<HTMLInputElement>(null);
  const latest = useRef(onPick);
  latest.current = onPick;
  useEffect(() => {
    const el = ref.current!;
    // "change" fires when the picker closes; "input" would fire for every hover.
    const onChange = () => latest.current(el.value);
    el.addEventListener("change", onChange);
    return () => el.removeEventListener("change", onChange);
  }, []);
  // Follow the active cell's colour without remounting the input.
  useEffect(() => {
    ref.current!.value = value;
  }, [value]);
  return (
    <label className="color-picker" title={title}>
      <span style={{ borderBottomColor: value }}>{label}</span>
      <input ref={ref} type="color" defaultValue={value} aria-label={title} />
    </label>
  );
}

function StatusBar({ doc, sheet, rect }: { doc: SheetsDocument; sheet: number; rect: Rect | null }) {
  if (!rect || (rect.top === rect.bottom && rect.left === rect.right)) return <div className="status-bar" />;
  const [count, numbers, sum] = doc.rangeStats(sheet, rect.top, rect.left, rect.bottom - rect.top + 1, rect.right - rect.left + 1);
  const fmt = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 10 });
  return (
    <div className="status-bar">
      {numbers > 0 && <span>Average: {fmt(sum / numbers)}</span>}
      <span>Count: {count}</span>
      {numbers > 0 && <span>Sum: {fmt(sum)}</span>}
    </div>
  );
}
