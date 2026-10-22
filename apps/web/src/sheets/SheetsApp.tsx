// Owns all spreadsheet UI state and turns user actions into engine calls.
// The engine (`SheetsDocument`, Rust via WebAssembly) owns the data; React
// state here is only selection, the edit in progress, menus, and a revision
// counter that tells the grid to repaint.

import { useEffect, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent } from "react";
import { SheetsDocument, ValueKind, functionList, parseStyle, type FunctionInfo } from "@office/engine";
import { Ribbon, type RibbonTab } from "@office/ui";
import { download, pickFile } from "../files";
import { saveDocument, touchDocument } from "../documents";
import type { Journal } from "./journal";
import { parseTsv, toTsv } from "./clipboard";
import { ContextMenu, type MenuItem } from "./ContextMenu";
import { FindBar, type FindOptions } from "./FindBar";
import { FormulaAssist } from "./FormulaAssist";
import { FormulaBar } from "./FormulaBar";
import { canInsertRef, closeParens, cycleReference, enclosingFunction, findRefs, refsOnSheet, wordAtCaret } from "./formulaText";
import {
  MAX_COLS,
  MAX_ROWS,
  ROW_HEIGHT,
  clampToExtent,
  moveSelection,
  rangeName,
  selectionRect,
  single,
  type Cell,
  type Rect,
  type Selection,
} from "./geometry";
import { alignmentCommands, bordersCommand, fontCommands, painterCommand, type BorderLineChoice, type FormattingContext, type PainterMode } from "./features/formatting";
import { SizeDialog, formatMenuItems, freezeCommand, mergeCommand, type Axis, type LayoutContext, type MergeKind } from "./features/layout";
import {
  CfGalleryDialog,
  CfManageDialog,
  CfRuleDialog,
  CustomFilterDialog,
  FilterMenu,
  ListDropdown,
  MessageDialog,
  RemoveDuplicatesDialog,
  SortDialog,
  Top10Dialog,
  ValidationAlert,
  ValidationDialog,
  conditionalFormattingCommand,
  dataTabCommands,
  readFilter,
  readValidation,
  sortFilterMenuItems,
  tableArea,
  type AlertChoice,
  type DataContext,
  type DataDialog,
  type ValidationRule,
} from "./features/dataTools";
import { navigate } from "./merges";
import { buildGeometry } from "./sheetGeometry";
import { Grid, fillArea, type FillDirection, type GridShortcut } from "./Grid";
import { fontFor } from "./paint";
import { NameBox } from "./NameBox";
import { SheetTabs } from "./SheetTabs";
import type { EditState } from "./types";
import { MOD_KEY, insertFunctionText, statusMode, zoomToFit } from "./chrome";
import { InsertFunctionDialog } from "./InsertFunctionDialog";
import { ColumnLayout, DEFAULT_COL_WIDTH } from "./layout";
import { StatusBar } from "./StatusBar";
import { WorkbookTitle, type SaveStatus } from "./WorkbookTitle";
import { Icon } from "@office/ui";
import { Select } from "./Select";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
/** Listed first in autocomplete when they match, since most people want these. */
const POPULAR = ["SUM", "AVERAGE", "COUNT", "IF", "VLOOKUP", "XLOOKUP", "SUMIF", "COUNTIF", "MAX", "MIN", "ROUND", "IFERROR", "TODAY", "CONCAT", "INDEX", "MATCH"];
const popularity = (name: string) => {
  const i = POPULAR.indexOf(name);
  return i === -1 ? POPULAR.length : i;
};

const ZOOM_KEY = "office.sheets.zoom";
const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2];

function savedZoom(): number {
  try {
    const z = Number(localStorage.getItem(ZOOM_KEY));
    return z >= 0.5 && z <= 2 ? z : 1;
  } catch {
    return 1;
  }
}

/** A full save quicker than this happens after every edit (see autosave). */
const CHEAP_FULL_SAVE_MS = 150;
/** Big workbooks are rewritten once this many edits are in the journal. */
const COMPACT_AFTER_EDITS = 200;

/** Formatting more cells than this at once is limited to the area holding data. */
const MAX_FORMAT_CELLS = 200_000;

interface SheetsAppProps {
  /** Id this workbook is autosaved under (see documents.ts). */
  docId: string;
  /** The opened workbook (see SheetsLoader.tsx). */
  initialDoc: SheetsDocument;
  /** Records the document's edits for autosave (see journal.ts). */
  journal: Journal;
  initialName: string;
  /** False when the workbook isn't stored yet (e.g. a CSV just opened). */
  initiallySaved: boolean;
  onNewWorkbook: () => void;
  onOpenFile: (file: File) => void;
}

export function SheetsApp({ docId, initialDoc, journal, initialName, initiallySaved, onNewWorkbook, onOpenFile }: SheetsAppProps) {
  const [doc, setDoc] = useState(initialDoc);
  const [fileName, setFileName] = useState(initialName);
  const [sheet, setSheet] = useState(0);
  const [selection, setSelection] = useState<Selection>(single({ row: 0, col: 0 }));
  const [editing, setEditing] = useState<EditState | null>(null);
  const [revision, setRevision] = useState(initiallySaved ? 0 : 1);
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
  // Excel chrome: Insert Function dialog, autosave status.
  const [insertFunctionOpen, setInsertFunctionOpen] = useState(false);
  /** The edit that was in progress when Insert Function opened; resumed when it closes. */
  const suspendedEdit = useRef<EditState | null>(null);
  const [savedRevision, setSavedRevision] = useState(0);
  const [saveFailed, setSaveFailed] = useState(false);
  const changed = () => setRevision((r) => r + 1);
  // Data tools: open dialog, filter dropdown, in-cell list, Circle Invalid Data, messages.
  const [dataDialog, setDataDialog] = useState<DataDialog | null>(null);
  const [filterMenu, setFilterMenu] = useState<{ col: number; x: number; y: number } | null>(null);
  const [listMenu, setListMenu] = useState<{ x: number; y: number; width: number } | null>(null);
  const [circlesOn, setCirclesOn] = useState(false);
  const [message, setMessage] = useState<{ title: string; message: string } | null>(null);
  /** An entry that broke a validation rule, waiting for the user's answer. */
  const [alertFor, setAlertFor] = useState<{ rule: ValidationRule; dRow: number; dCol: number } | null>(null);

  const { row, col } = selection.anchor;
  const sel = selectionRect(selection);
  const activeStyle = parseStyle(doc.styleBlock(sheet, row, col, 1, 1)[0]);
  // Row heights (incl. automatic ones), column widths, hidden lines, merges, frozen panes.
  const geo = useMemo(() => buildGeometry(doc, sheet), [doc, sheet, revision]);
  const filter = useMemo(() => readFilter(doc, sheet), [doc, sheet, revision]);
  const circles = useMemo(() => {
    if (!circlesOn) return [];
    const flat = doc.invalidCells(sheet, 0, 0, MAX_ROWS - 1, MAX_COLS - 1);
    const out: Cell[] = [];
    for (let i = 0; i < flat.length; i += 2) out.push({ row: flat[i], col: flat[i + 1] });
    return out;
  }, [doc, sheet, revision, circlesOn]);
  const activeRule = readValidation(doc.validationAt(sheet, row, col));
  const listItems = activeRule?.kind === "list" && activeRule.dropdown ? doc.listItems(sheet, row, col) : [];
  const inputMessage =
    activeRule?.showInput && (activeRule.inputTitle || activeRule.inputMessage) ? { title: activeRule.inputTitle, message: activeRule.inputMessage } : null;

  /** Select and scroll into view. */
  /**
   * Select (snapping to merged cells) and scroll into view. Go To and Find show
   * the start of the range, like Excel; Ctrl+Shift+arrow shows the moving end.
   */
  function goTo(s: Selection, show: "anchor" | "focus" = "anchor") {
    const next = geo.merges.normalize(s);
    setSelection(next);
    const cell = show === "anchor" ? next.anchor : next.focus;
    setReveal((r) => ({ row: cell.row, col: cell.col, seq: (r?.seq ?? 0) + 1 }));
  }

  function setZoom(z: number) {
    setZoomState(z);
    try {
      localStorage.setItem(ZOOM_KEY, String(z));
    } catch {
      // ignore
    }
  }

  // ---- autosave
  //
  // Every edit is appended to a journal right away (journal.ts): a few
  // milliseconds whatever the workbook's size. The full .xlsx is rewritten
  // ("compacted") after each edit when that is cheap; for big workbooks only
  // once many edits have piled up and the user pauses, so typing never stalls.

  /** How long the last full save took; null until the first one. */
  const fullSaveMs = useRef<number | null>(null);
  /** The stored file doesn't hold this workbook yet (e.g. a CSV just opened). */
  const needsFullSave = useRef(!initiallySaved);

  function estimatedSaveMs(): number {
    if (fullSaveMs.current !== null) return fullSaveMs.current;
    // About a microsecond per cell in the WebAssembly build.
    return doc.sheetNames().reduce((sum, _, i) => {
      const [rows, cols] = doc.extent(i);
      return sum + rows * cols * 0.001;
    }, 0);
  }

  function fullSave() {
    const covered = journal.length;
    const started = performance.now();
    const bytes = new Uint8Array(doc.toXlsx());
    fullSaveMs.current = performance.now() - started;
    needsFullSave.current = false;
    void saveDocument(docId, "sheets", fileName, bytes).then((ok) => {
      if (ok) void journal.dropFirst(covered);
      else setSaveFailed(true);
    });
  }

  useEffect(() => {
    if (revision === savedRevision) return;
    const rev = revision;
    const timer = setTimeout(() => {
      touchDocument(docId, "sheets", fileName);
      void journal.flush().then((ok) => {
        setSaveFailed(!ok);
        setSavedRevision((r) => Math.max(r, rev));
      });
      if (needsFullSave.current || estimatedSaveMs() < CHEAP_FULL_SAVE_MS) fullSave();
    }, 300);
    return () => clearTimeout(timer);
  }, [doc, revision]);

  // Big workbooks: rewrite the file after a pause once enough edits piled up.
  useEffect(() => {
    if (journal.length < COMPACT_AFTER_EDITS) return;
    const timer = setTimeout(fullSave, 10_000);
    return () => clearTimeout(timer);
  }, [revision]);

  // A rename only updates the name shown in Recent.
  const firstName = useRef(fileName);
  useEffect(() => {
    if (fileName !== firstName.current) touchDocument(docId, "sheets", fileName);
  }, [fileName]);

  // Leaving (tab hidden, page closed, Home): store the journal right away.
  const flush = useRef(() => {});
  flush.current = () => {
    if (revision !== savedRevision) void journal.flush();
  };
  useEffect(() => {
    const onHide = () => document.visibilityState === "hidden" && flush.current();
    const onLeave = () => flush.current();
    document.addEventListener("visibilitychange", onHide);
    window.addEventListener("pagehide", onLeave);
    return () => {
      document.removeEventListener("visibilitychange", onHide);
      window.removeEventListener("pagehide", onLeave);
      flush.current();
    };
  }, []);

  useEffect(() => {
    document.title = `${fileName} - Sheets`;
  }, [fileName]);

  const saveStatus: SaveStatus = revision !== savedRevision ? "saving" : saveFailed ? "failed" : "saved";

  // ---- editing

  function startEdit(source: EditState["source"], initialText?: string) {
    const draft = initialText ?? doc.input(sheet, row, col);
    setEditing({ draft, source, mode: initialText === undefined ? "edit" : "enter", caret: draft.length, sheet, cell: { row, col } });
    setAssistDismissedAt(null);
    setAssistIndex(0);
  }

  function changeDraft(draft: string, caret: number) {
    setEditing((e) => e && { ...e, draft, caret, point: undefined });
    setAssistIndex(0);
  }

  /** Where a run of Tab-committed entries started (see commit). */
  const tabStart = useRef<Cell | null>(null);

  /** `force`: the user already accepted an entry that breaks a Warning/Information rule. */
  function commit(dRow: number, dCol: number, force = false) {
    if (!editing) return;
    // The entry goes where the edit started, even if another sheet is showing.
    const target = editing.sheet;
    const { row: r, col: c } = editing.cell;
    const value = closeParens(editing.draft);
    if (value !== doc.input(target, r, c)) {
      const broken = force ? null : readValidation(doc.checkInput(target, r, c, value));
      if (broken) {
        setAlertFor({ rule: broken, dRow, dCol });
        return;
      }
      // A line break (Alt+Enter) turns on Wrap Text in the same undo step, like Excel.
      const wrapped = !!parseStyle(doc.styleBlock(target, r, c, 1, 1)[0]).wrap;
      if (value.includes("\n") && !wrapped) doc.setInputWrapped(target, r, c, value);
      else doc.setInputs(target, Uint32Array.of(r), Uint32Array.of(c), [value]);
      widenForNumber(r, c, target);
      changed();
    }
    setEditing(null);
    // Excel: Enter after a run of Tab entries goes to the next row of the column where the run started.
    const start = tabStart.current;
    tabStart.current = dCol !== 0 && dRow === 0 ? (start?.row === r ? start : { row: r, col: c }) : null;
    if (target !== sheet) {
      setSheet(target);
      setSelection(single(editing.cell));
    } else if (dRow === 1 && dCol === 0 && start?.row === r && selection.anchor.row === r) {
      setSelection(single({ row: geo.rows.nextVisible(r + 1, 1) ?? r + 1, col: start.col }));
    } else if (dRow || dCol) {
      setSelection(navigate(selection, dRow, dCol, false, geo));
    }
  }

  /**
   * Like Excel: when a number or date typed into a default-width column
   * doesn't fit, the column widens instead of showing ####.
   */
  function widenForNumber(r: number, c: number, s = sheet) {
    if (doc.kindBlock(s, r, c, 1, 1)[0] !== ValueKind.Number) return;
    const widths = doc.columnWidths(s);
    for (let i = 0; i < widths.length; i += 2) if (widths[i] === c) return; // user-set width: keep it
    const ctx = document.createElement("canvas").getContext("2d")!;
    ctx.font = fontFor(parseStyle(doc.styleBlock(sheet, r, c, 1, 1)[0]));
    const needed = Math.ceil(ctx.measureText(doc.displayBlock(sheet, r, c, 1, 1)[0]).width + 10);
    if (needed > DEFAULT_COL_WIDTH) doc.setColumnWidth(sheet, c, c, needed);
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
  const highlights = useMemo(() => {
    if (!isFormula) return [];
    const names = doc.sheetNames();
    return refsOnSheet(findRefs(editing!.draft), names[sheet], names[editing!.sheet]);
  }, [isFormula, editing?.draft, editing?.sheet, sheet, revision]);
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
    if (e.key === "F4" && isFormula) {
      // F4: cycle $A$1 -> A$1 -> $A1 -> A1 for the reference at the caret.
      e.preventDefault();
      const next = cycleReference(editing.draft, editing.caret);
      if (next) replaceDraft(next.formula, next.caret);
      return true;
    }
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
    // Pointing on another sheet writes Sheet2!A1 (quoted when the name needs it).
    const prefix = sheet === editing.sheet ? "" : `${SheetsDocument.quoteSheetName(doc.sheetNames()[sheet])}!`;
    const text = prefix + rangeName(s);
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

  // ---- borders and Format Painter

  const [borderLine, setBorderLine] = useState<BorderLineChoice>({ style: "thin", color: "" });
  const [lastBorder, setLastBorder] = useState("bottom");
  const [painter, setPainterMode] = useState<PainterMode>("off");

  function applyBorders(preset: string) {
    let r: Rect | null = sel;
    if ((r.bottom - r.top + 1) * (r.right - r.left + 1) > MAX_FORMAT_CELLS) r = dataRect();
    if (!r) return;
    doc.applyBorders(sheet, r.top, r.left, r.bottom, r.right, preset, borderLine.style, borderLine.color);
    setLastBorder(preset);
    changed();
  }

  function setPainter(mode: PainterMode) {
    if (mode !== "off") doc.pickFormats(sheet, sel.top, sel.left, sel.bottom, sel.right);
    setPainterMode(mode);
  }

  // While Format Painter is on, the next selection made in the grid receives the
  // formatting when the mouse is released (Excel behaviour). Esc stops it.
  const selectionRef = useRef(selection);
  selectionRef.current = selection;
  useEffect(() => {
    if (painter === "off") return;
    const onUp = (e: MouseEvent) => {
      if (!(e.target as Element | null)?.closest?.(".grid")) return;
      const r = selectionRect(selectionRef.current);
      if (doc.paintFormats(sheet, r.top, r.left, r.bottom, r.right).length) changed();
      if (painter === "once") setPainterMode("off");
    };
    const onKey = (e: globalThis.KeyboardEvent) => e.key === "Escape" && setPainterMode("off");
    window.addEventListener("mouseup", onUp);
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("mouseup", onUp);
      window.removeEventListener("keydown", onKey);
    };
  }, [painter, doc, sheet]);

  const formatting: FormattingContext = {
    style: activeStyle,
    format,
    applyBorders,
    borderLine,
    setBorderLine,
    lastBorder,
    painter,
    setPainter,
    focusGrid: () => cellInputRef.current?.focus({ preventScroll: true }),
  };

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

  // ---- rows: heights, autofit; hide/unhide; merges; freeze panes

  const wholeRows = sel.left === 0 && sel.right === MAX_COLS - 1;
  const wholeCols = sel.top === 0 && sel.bottom === MAX_ROWS - 1;
  const [sizeDialog, setSizeDialog] = useState<Axis | null>(null);

  function resizeRow(r: number, px: number) {
    // Resizing one of several selected whole rows resizes them all.
    const many = wholeRows && r >= sel.top && r <= sel.bottom;
    doc.setRowHeight(sheet, many ? sel.top : r, many ? sel.bottom : r, px);
    changed();
  }

  /** Back to automatic height, which grows to fit wrapped text and big fonts. */
  function autofitRow(r: number) {
    const many = wholeRows && r >= sel.top && r <= sel.bottom;
    doc.setRowHeight(sheet, many ? sel.top : r, many ? sel.bottom : r, -1);
    changed();
  }

  /** Rows or columns the selection covers, limited to where there is data for whole-sheet selections. */
  function lineSpanOf(axis: Axis): [number, number] {
    const [rows, cols] = doc.extent(sheet);
    if (axis === "rows") return [sel.top, Math.min(sel.bottom, wholeCols ? Math.max(rows, sel.top + 1) - 1 : sel.bottom)];
    return [sel.left, Math.min(sel.right, wholeRows ? Math.max(cols, sel.left + 1) - 1 : sel.right)];
  }

  const layoutCtx: LayoutContext = {
    sel,
    merges: geo.merges,
    freeze: geo.freeze,
    active: selection.anchor,
    merge(kind: MergeKind) {
      if (doc.mergeDiscardsData(sheet, sel.top, sel.left, sel.bottom, sel.right, kind)) {
        if (!confirm("Merging cells only keeps the upper-left value and discards other values.")) return;
      }
      doc.mergeCells(sheet, sel.top, sel.left, sel.bottom, sel.right, kind);
      setSelection(single(selection.anchor));
      changed();
    },
    unmerge() {
      doc.unmergeCells(sheet, sel.top, sel.left, sel.bottom, sel.right);
      changed();
    },
    setFreeze(rows: number, cols: number) {
      doc.setFreezePanes(sheet, rows, cols);
      changed();
    },
    setHidden(axis: Axis, hidden: boolean) {
      const [first, last] = lineSpanOf(axis);
      doc.setHidden(sheet, axis, first, last, hidden);
      changed();
    },
    openSizeDialog: setSizeDialog,
    autofit(axis: Axis) {
      const [first, last] = lineSpanOf(axis);
      if (axis === "rows") doc.setRowHeight(sheet, first, last, -1);
      else for (let c = first; c <= last; c++) autofit(c);
      changed();
    },
  };

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
    if (!doc.canSort(sheet, r.top, r.left, r.bottom, r.right)) {
      alert("To do this, all the merged cells need to be the same size.");
      return;
    }
    doc.sortRange(sheet, r.top, r.left, r.bottom, r.right, col, ascending);
    changed();
  }

  // ---- data tools: filter, sort, remove duplicates, validation, conditional formatting

  function toggleFilter() {
    const area = tableArea(doc, sheet, sel);
    doc.toggleFilter(sheet, area.top, area.left, area.bottom, area.right);
    setFilterMenu(null);
    changed();
  }

  function answerAlert(choice: AlertChoice) {
    const pending = alertFor;
    setAlertFor(null);
    if (!pending) return;
    if (choice === "accept") commit(pending.dRow, pending.dCol, true);
    else if (choice === "cancel") setEditing(null);
    else {
      // Retry: back to the entry with its text selected, like Excel.
      requestAnimationFrame(() => {
        const el = editing?.source === "bar" ? barInputRef.current : cellInputRef.current;
        el?.focus({ preventScroll: true });
        el?.select();
      });
    }
  }

  /** Alt+Down: the active cell's list, or its column's filter dropdown. */
  function openDropdown() {
    const box = cellInputRef.current?.getBoundingClientRect();
    if (!box) return;
    if (listItems.length) setListMenu({ x: box.left, y: box.bottom + 1, width: box.width });
    else if (filter && row === filter.range.top && col >= filter.range.left && col <= filter.range.right)
      setFilterMenu({ col, x: box.left, y: box.bottom + 1 });
  }

  function pickFromList(item: string) {
    setListMenu(null);
    doc.setInputs(sheet, Uint32Array.of(row), Uint32Array.of(col), [item]);
    changed();
    cellInputRef.current?.focus({ preventScroll: true });
  }

  const dataCtx: DataContext = {
    filter,
    toggleFilter,
    clearFilters() {
      doc.clearFilters(sheet);
      changed();
    },
    reapplyFilter() {
      doc.reapplyFilter(sheet);
      changed();
    },
    sort,
    openDialog: setDataDialog,
    clearCf(whole: boolean) {
      doc.clearCfRules(sheet, whole, sel.top, sel.left, sel.bottom, sel.right);
      changed();
    },
    circles: circlesOn,
    setCircles: setCirclesOn,
  };

  const filterCounts = filter ? doc.filterCounts(sheet) : null;
  const filterStatus = filterCounts && filterCounts.length === 2 ? `${filterCounts[0]} of ${filterCounts[1]} records found` : "";

  function closeDataDialog() {
    setDataDialog(null);
    cellInputRef.current?.focus({ preventScroll: true });
  }

  function renderDataDialog() {
    const d = dataDialog;
    if (!d) return null;
    const close = closeDataDialog;
    switch (d.kind) {
      case "sort":
        return (
          <SortDialog
            doc={doc}
            sheet={sheet}
            sel={sel}
            activeCol={col}
            onDone={(area) => {
              changed();
              setSelection({ anchor: { row: area.top, col: area.left }, focus: { row: area.bottom, col: area.right } });
            }}
            onClose={close}
          />
        );
      case "dedupe":
        return (
          <RemoveDuplicatesDialog
            doc={doc}
            sheet={sheet}
            sel={sel}
            onDone={(area, text) => {
              changed();
              setSelection({ anchor: { row: area.top, col: area.left }, focus: { row: area.bottom, col: area.right } });
              setMessage({ title: "Remove Duplicates", message: text });
            }}
            onClose={close}
          />
        );
      case "validation":
        return <ValidationDialog doc={doc} sheet={sheet} sel={sel} active={selection.anchor} onDone={changed} onClose={close} />;
      case "customFilter":
        return filter && <CustomFilterDialog doc={doc} sheet={sheet} filter={filter} col={d.col} preset={d.preset} onDone={changed} onClose={close} />;
      case "top10":
        return <Top10Dialog doc={doc} sheet={sheet} col={d.col} onDone={changed} onClose={close} />;
      case "cf":
        return <CfRuleDialog doc={doc} sheet={sheet} sel={sel} rule={d.rule} onDone={changed} onClose={close} />;
      case "cfGallery":
        return <CfGalleryDialog doc={doc} sheet={sheet} sel={sel} gallery={d.gallery} onDone={changed} onClose={close} />;
      case "cfManage":
        return <CfManageDialog doc={doc} sheet={sheet} onDone={changed} onClose={close} />;
    }
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
    goTo(extend ? { anchor: selection.anchor, focus: to } : single(to), extend ? "focus" : "anchor");
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
      filter: toggleFilter,
      dropdown: openDropdown,
    };
    actions[s]();
  }

  function undo() {
    if (doc.undo()) changed();
  }

  function redo() {
    if (doc.redo()) changed();
  }

  // ---- Excel chrome: Insert Function, zoom to selection, chrome shortcuts

  /** fx / Shift+F3. An edit in progress is set aside (not committed) while the dialog is open. */
  function openInsertFunction() {
    suspendedEdit.current = editing;
    setEditing(null);
    setInsertFunctionOpen(true);
  }

  function closeInsertFunction(f: FunctionInfo | null) {
    setInsertFunctionOpen(false);
    const before = suspendedEdit.current;
    suspendedEdit.current = null;
    if (!f) {
      if (before) {
        setEditing(before);
        pendingCaret.current = before.caret;
      }
      return;
    }
    const next = insertFunctionText(before?.draft ?? null, before?.caret ?? 0, f.name);
    setEditing({
      draft: next.draft,
      caret: next.caret,
      source: before?.source ?? "cell",
      mode: "edit",
      sheet: before?.sheet ?? sheet,
      cell: before?.cell ?? { row, col },
    });
    pendingCaret.current = next.caret;
    setAssistDismissedAt(null);
    setAssistIndex(0);
  }

  function zoomToSelection() {
    const el = document.querySelector<HTMLElement>(".grid-scroller");
    if (!el) return;
    const layout = new ColumnLayout(doc.columnWidths(sheet));
    const width = layout.x(sel.right + 1) - layout.x(sel.left);
    const height = (sel.bottom - sel.top + 1) * ROW_HEIGHT;
    setZoom(zoomToFit(width, height, el.clientWidth, el.clientHeight));
    goTo({ anchor: { row: sel.top, col: sel.left }, focus: { row: sel.bottom, col: sel.right } });
  }

  const chromeKeys = useRef<(e: globalThis.KeyboardEvent) => void>(() => {});
  chromeKeys.current = (e) => {
    if (e.key === "F3" && e.shiftKey && !insertFunctionOpen) {
      e.preventDefault();
      openInsertFunction();
    } else if (e.key === "F11" && e.shiftKey) {
      e.preventDefault();
      if (editing) commit(0, 0);
      addSheet();
    }
  };
  useEffect(() => {
    const onKey = (e: globalThis.KeyboardEvent) => chromeKeys.current(e);
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // ---- files

  async function open() {
    const file = await pickFile(".xlsx,.csv");
    if (file) onOpenFile(file);
  }

  function newDocument() {
    // The current workbook is autosaved and stays in Recent on Home.
    flush.current();
    onNewWorkbook();
  }

  // ---- menus

  function cellMenu(x: number, y: number) {
    const mod = navigator.platform.startsWith("Mac") ? "⌘" : "Ctrl+";
    const items: MenuItem[] = [
      { label: "Cut", shortcut: `${mod}X`, onRun: () => navigator.clipboard?.writeText(cut()), icon: "cut" },
      { label: "Copy", shortcut: `${mod}C`, onRun: () => navigator.clipboard?.writeText(copy()), icon: "copy" },
      { label: "Paste", shortcut: `${mod}V`, onRun: pasteFromMenu, icon: "paste" },
      { label: "Paste values only", shortcut: `${mod}⇧V`, onRun: () => pasteSpecial("values", false), disabled: !doc.hasCopied(), icon: "pasteSpecial" },
      { label: "Paste formatting only", onRun: () => pasteSpecial("formats", false), disabled: !doc.hasCopied(), icon: "formatPainter" },
      { label: "Paste formulas only", onRun: () => pasteSpecial("formulas", false), disabled: !doc.hasCopied(), icon: "function" },
      { label: "Paste transposed", onRun: () => pasteSpecial("all", true), disabled: !doc.hasCopied() },
      { label: "Insert row above", onRun: () => insertLines("rows", false), separatorBefore: true, icon: "insertRowAbove" },
      { label: "Insert row below", onRun: () => insertLines("rows", true), icon: "insertRowBelow" },
      { label: "Insert column left", onRun: () => insertLines("cols", false), icon: "insertColumnLeft" },
      { label: "Insert column right", onRun: () => insertLines("cols", true), icon: "insertColumnRight" },
      { label: "Delete rows", onRun: () => deleteLines("rows"), separatorBefore: true, icon: "deleteRows" },
      { label: "Delete columns", onRun: () => deleteLines("cols"), icon: "deleteColumns" },
      ...(wholeRows
        ? [
            { label: "Row Height…", onRun: () => setSizeDialog("rows"), separatorBefore: true },
            { label: "Hide", onRun: () => layoutCtx.setHidden("rows", true) },
            { label: "Unhide", onRun: () => layoutCtx.setHidden("rows", false) },
          ]
        : []),
      ...(wholeCols
        ? [
            { label: "Column Width…", onRun: () => setSizeDialog("cols"), separatorBefore: true },
            { label: "Hide", onRun: () => layoutCtx.setHidden("cols", true) },
            { label: "Unhide", onRun: () => layoutCtx.setHidden("cols", false) },
          ]
        : []),
      { label: "Sort A → Z", onRun: () => sort(true), separatorBefore: true, icon: "sortAsc" },
      { label: "Sort Z → A", onRun: () => sort(false), icon: "sortDesc" },
      { label: "Clear contents", shortcut: "Del", onRun: clear, separatorBefore: true, icon: "clear" },
      { label: "Clear formatting", onRun: () => format("clear", "") },
      { label: "Insert Function…", shortcut: "⇧F3", onRun: openInsertFunction, separatorBefore: true, icon: "function" },
    ];
    setMenu({ x, y, items });
  }

  function tabMenu(index: number, x: number, y: number) {
    setMenu({
      x,
      y,
      items: [
        { label: "Rename", onRun: () => setRenamingTab(index), icon: "rename" },
        {
          label: "Delete sheet",
          icon: "trash",
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
  // Tabs and groups follow Excel's Home / Insert / Formulas / Data / View
  // layout; see CONTRIBUTING.md ("Add a ribbon command") for group names.
  const copied = doc.hasCopied();
  const tabs: RibbonTab[] = [
    {
      id: "file",
      label: "File",
      commands: [
        { id: "new", label: "New", title: "Start a new, empty workbook.", icon: "newDocument", size: "large", group: "Workbook", onRun: newDocument },
        { id: "open", label: "Open…", title: "Open an .xlsx or .csv file from your computer.", icon: "open", size: "large", group: "Workbook", onRun: open },
        {
          id: "save-xlsx",
          label: "Download .xlsx",
          title: "Save a copy of the workbook as an Excel file.",
          icon: "excelFile",
          size: "large",
          group: "Save a Copy",
          onRun: () => download(new Uint8Array(doc.toXlsx()), `${fileName}.xlsx`, XLSX_MIME),
        },
        {
          id: "save-csv",
          label: "Download .csv",
          title: "Save the current sheet as comma-separated values.",
          icon: "csvFile",
          size: "large",
          group: "Save a Copy",
          onRun: () => download(doc.toCsv(sheet), `${fileName}.csv`, "text/csv"),
        },
      ],
    },
    {
      id: "home",
      label: "Home",
      commands: [
        { id: "undo", label: "Undo", shortcut: `${MOD_KEY}Z`, icon: "undo", size: "small", group: "Undo", onRun: undo },
        { id: "redo", label: "Redo", shortcut: `${MOD_KEY}Y`, icon: "redo", size: "small", group: "Undo", onRun: redo },
        {
          id: "paste",
          label: "Paste",
          shortcut: `${MOD_KEY}V`,
          title: "Add content from the clipboard to the selected cells.",
          icon: "paste",
          size: "large",
          group: "Clipboard",
          onRun: pasteFromMenu,
          menu: [
            { id: "paste-all", label: "Paste", shortcut: `${MOD_KEY}V`, icon: "paste", onRun: pasteFromMenu },
            { id: "paste-values", label: "Values", shortcut: `${MOD_KEY}⇧V`, icon: "pasteSpecial", disabled: !copied, onRun: () => pasteSpecial("values", false), separatorBefore: true },
            { id: "paste-formulas", label: "Formulas", icon: "function", disabled: !copied, onRun: () => pasteSpecial("formulas", false) },
            { id: "paste-formats", label: "Formatting", icon: "formatPainter", disabled: !copied, onRun: () => pasteSpecial("formats", false) },
            { id: "paste-transpose", label: "Transpose", icon: "table", disabled: !copied, onRun: () => pasteSpecial("all", true) },
          ],
        },
        { id: "cut", label: "Cut", shortcut: `${MOD_KEY}X`, icon: "cut", size: "small", group: "Clipboard", onRun: () => navigator.clipboard?.writeText(cut()) },
        { id: "copy", label: "Copy", shortcut: `${MOD_KEY}C`, icon: "copy", size: "small", group: "Clipboard", onRun: () => navigator.clipboard?.writeText(copy()) },
        painterCommand(formatting),
        ...fontCommands(formatting),
        { id: "bold", newRow: true, label: "Bold", shortcut: `${MOD_KEY}B`, title: "Make your text bold.", icon: "bold", group: "Font", active: !!activeStyle.b, onRun: () => toggle("bold", !!activeStyle.b) },
        { id: "italic", label: "Italic", shortcut: `${MOD_KEY}I`, title: "Italicize your text.", icon: "italic", group: "Font", active: !!activeStyle.i, onRun: () => toggle("italic", !!activeStyle.i) },
        { id: "underline", label: "Underline", shortcut: `${MOD_KEY}U`, title: "Underline your text.", icon: "underline", group: "Font", active: !!activeStyle.u, onRun: () => toggle("underline", !!activeStyle.u) },
        { id: "strike", label: "Strikethrough", title: "Cross out text by drawing a line through it.", icon: "strike", group: "Font", active: !!activeStyle.s, onRun: () => toggle("strike", !!activeStyle.s) },
        bordersCommand(formatting),
        {
          id: "fill",
          label: "Fill Color",
          title: "Color the background of cells to make them stand out.",
          group: "Font",
          render: () => <ColorPicker label="▧" icon="fillColor" title="Fill colour" value={activeStyle.fill ?? "#ffffff"} onPick={(c) => format("fill", c)} />,
        },
        {
          id: "color",
          label: "Font Color",
          title: "Change the color of your text.",
          group: "Font",
          render: () => <ColorPicker label="A" icon="fontColor" title="Text colour" value={activeStyle.color ?? "#1f1f1f"} onPick={(c) => format("color", c)} />,
        },
        ...alignmentCommands(formatting),
        { id: "align-left", newRow: true, label: "Align Left", title: "Align your content to the left.", icon: "alignLeft", group: "Alignment", active: activeStyle.align === "left", onRun: () => format("align", "left") },
        { id: "align-center", label: "Center", title: "Center your content.", icon: "alignCenter", group: "Alignment", active: activeStyle.align === "center", onRun: () => format("align", "center") },
        { id: "align-right", label: "Align Right", title: "Align your content to the right.", icon: "alignRight", group: "Alignment", active: activeStyle.align === "right", onRun: () => format("align", "right") },
        mergeCommand(layoutCtx),
        {
          id: "format",
          label: "Number Format",
          title: "Choose how values in your cells are displayed: as a percentage, currency, date or time, etc.",
          group: "Number",
          render: () => (
            <Select className="number-format" aria-label="Number format" value={formatKind} onChange={(e) => format("format", e.target.value)}>
              <option value="general">General</option>
              <option value="number">Number (1,234.50)</option>
              <option value="currency">Rupee (₹1,234.50)</option>
              <option value="dollar">Dollar ($1,234.50)</option>
              <option value="percent">Percent (12%)</option>
              <option value="date">Date (15-08-2025)</option>
              <option value="time">Time (14:30)</option>
              <option value="datetime">Date &amp; time</option>
              <option value="text">Text</option>
            </Select>
          ),
        },
        { id: "fmt-currency", label: "Currency Format", title: "Format as Indian rupees (₹1,234.50).", icon: "currency", group: "Number", newRow: true, active: formatKind === "currency", onRun: () => format("format", "currency") },
        { id: "fmt-percent", label: "Percent Style", title: "Format as a percent.", icon: "percent", group: "Number", active: formatKind === "percent", onRun: () => format("format", "percent") },
        { id: "fmt-comma", label: "Comma Style", title: "Format with a thousands separator (1,234.50).", icon: "comma", group: "Number", active: formatKind === "number", onRun: () => format("format", "number") },
        { id: "dec-more", label: "Increase Decimal", title: "Show more decimal places for a more precise value.", icon: "decimalIncrease", group: "Number", separatorBefore: true, onRun: () => format("decimals", "1") },
        { id: "dec-less", label: "Decrease Decimal", title: "Show fewer decimal places.", icon: "decimalDecrease", group: "Number", onRun: () => format("decimals", "-1") },
        conditionalFormattingCommand(dataCtx),
        {
          id: "cells-insert",
          label: "Insert",
          title: "Add cells, rows, columns or sheets to your workbook.",
          icon: "insertCells",
          size: "small",
          group: "Cells",
          menu: [
            { id: "menu-row-above", label: "Insert Sheet Rows", icon: "insertRowAbove", onRun: () => insertLines("rows", false) },
            { id: "menu-row-below", label: "Insert Sheet Rows Below", icon: "insertRowBelow", onRun: () => insertLines("rows", true) },
            { id: "menu-col-left", label: "Insert Sheet Columns", icon: "insertColumnLeft", onRun: () => insertLines("cols", false), separatorBefore: true },
            { id: "menu-col-right", label: "Insert Sheet Columns Right", icon: "insertColumnRight", onRun: () => insertLines("cols", true) },
            { id: "menu-sheet", label: "Insert Sheet", shortcut: "⇧F11", icon: "insertSheet", onRun: addSheet, separatorBefore: true },
          ],
        },
        {
          id: "cells-delete",
          label: "Delete",
          title: "Delete rows or columns from your sheet.",
          icon: "deleteCells",
          size: "small",
          group: "Cells",
          menu: [
            { id: "menu-del-rows", label: "Delete Sheet Rows", icon: "deleteRows", onRun: () => deleteLines("rows") },
            { id: "menu-del-cols", label: "Delete Sheet Columns", icon: "deleteColumns", onRun: () => deleteLines("cols") },
          ],
        },
        {
          id: "cells-format",
          label: "Format",
          title: "Change row height or column width, hide rows or columns, or rename the sheet.",
          icon: "formatCells",
          size: "small",
          group: "Cells",
          menu: [
            ...formatMenuItems(layoutCtx),
            { id: "menu-rename-sheet", label: "Rename Sheet", icon: "rename", onRun: () => setRenamingTab(sheet), separatorBefore: true },
          ],
        },
        {
          id: "clear",
          label: "Clear",
          title: "Delete everything in the cell, or remove just the formatting or contents.",
          icon: "clear",
          size: "small",
          group: "Editing",
          menu: [
            { id: "clear-contents", label: "Clear Contents", shortcut: "Del", onRun: clear },
            { id: "clear-format", label: "Clear Formats", onRun: () => format("clear", "") },
          ],
        },
        {
          id: "sort-filter",
          label: "Sort & Filter",
          title: "Organize your data so it's easier to analyze.",
          icon: "sortFilter",
          size: "large",
          group: "Editing",
          menu: [
            { id: "sort-asc", label: "Sort A to Z", icon: "sortAsc", onRun: () => sort(true) },
            { id: "sort-desc", label: "Sort Z to A", icon: "sortDesc", onRun: () => sort(false) },
            ...sortFilterMenuItems(dataCtx),
          ],
        },
        {
          id: "find-select",
          label: "Find & Select",
          title: "Find and select specific text, formatting or type of information within the workbook.",
          icon: "find",
          size: "large",
          group: "Editing",
          menu: [
            { id: "find", label: "Find…", shortcut: `${MOD_KEY}F`, icon: "find", onRun: () => setFindMode("find") },
            { id: "replace", label: "Replace…", shortcut: `${MOD_KEY}H`, icon: "replace", onRun: () => setFindMode("replace") },
            { id: "goto", label: "Go To…", shortcut: `${MOD_KEY}G`, icon: "goTo", onRun: () => nameBoxRef.current?.focus() },
          ],
        },
      ],
    },
    {
      id: "insert",
      label: "Insert",
      commands: [
        { id: "row-above", label: "Row above", title: "Insert rows above the selection.", icon: "insertRowAbove", size: "small", group: "Cells", onRun: () => insertLines("rows", false) },
        { id: "row-below", label: "Row below", title: "Insert rows below the selection.", icon: "insertRowBelow", size: "small", group: "Cells", onRun: () => insertLines("rows", true) },
        { id: "col-left", label: "Column left", title: "Insert columns left of the selection.", icon: "insertColumnLeft", size: "small", group: "Cells", onRun: () => insertLines("cols", false) },
        { id: "col-right", label: "Column right", title: "Insert columns right of the selection.", icon: "insertColumnRight", size: "small", group: "Cells", onRun: () => insertLines("cols", true) },
        { id: "del-rows", label: "Delete rows", title: "Delete the selected rows.", icon: "deleteRows", size: "small", group: "Cells", onRun: () => deleteLines("rows") },
        { id: "del-cols", label: "Delete columns", title: "Delete the selected columns.", icon: "deleteColumns", size: "small", group: "Cells", onRun: () => deleteLines("cols") },
        { id: "sheet", label: "New sheet", shortcut: "⇧F11", title: "Add a sheet after the last one.", icon: "insertSheet", size: "large", group: "Sheets", onRun: addSheet },
      ],
    },
    {
      id: "formulas",
      label: "Formulas",
      commands: [
        {
          id: "insert-function",
          label: "Insert Function",
          shortcut: "⇧F3",
          title: "Search for a function and insert it into the active cell.",
          icon: "function",
          size: "large",
          group: "Function Library",
          onRun: openInsertFunction,
        },
      ],
    },
    {
      id: "data",
      label: "Data",
      commands: dataTabCommands(dataCtx),
    },
    {
      id: "view",
      label: "View",
      commands: [
        {
          id: "zoom",
          label: "Zoom",
          title: "Zoom in or out of the sheet.",
          icon: "zoom",
          size: "large",
          group: "Zoom",
          menu: ZOOM_LEVELS.map((z) => ({ id: `zoom-${z * 100}`, label: `${Math.round(z * 100)}%`, active: z === zoom, onRun: () => setZoom(z) })),
        },
        { id: "zoom-100", label: "100%", title: "Zoom to 100% of the normal size.", icon: "zoom100", size: "large", group: "Zoom", onRun: () => setZoom(1) },
        freezeCommand(layoutCtx),
        { id: "zoom-selection", label: "Zoom to Selection", title: "Zoom so the selected cells fill the window.", icon: "zoomSelection", size: "large", group: "Zoom", onRun: zoomToSelection },
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
      <WorkbookTitle
        name={fileName}
        onRename={(name) => {
          setFileName(name);
          changed();
        }}
        onDone={() => cellInputRef.current?.focus({ preventScroll: true })}
        status={saveStatus}
        onSave={() => download(new Uint8Array(doc.toXlsx()), `${fileName}.xlsx`, XLSX_MIME)}
        onUndo={undo}
        onRedo={redo}
      />
      <Ribbon tabs={tabs} initialTab="home" />
      <FormulaBar
        inputRef={barInputRef}
        nameBox={
          <NameBox
            ref={nameBoxRef}
            label={geo.merges.label(selection)}
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
        onInsertFunction={openInsertFunction}
      />
      <div className={painter === "off" ? "grid-area" : "grid-area painting"}>
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
          onResizeRow={resizeRow}
          onAutofitRow={autofitRow}
          geo={geo}
          onAutofit={autofit}
          onContextMenu={cellMenu}
          onShortcut={shortcut}
          filter={filter}
          onFilterButton={(c, x, y) => setFilterMenu((m) => (m?.col === c ? null : { col: c, x, y }))}
          circles={circles}
          listArrow={listItems.length > 0}
          onListArrow={(x, y, width) => setListMenu((m) => (m ? null : { x, y, width }))}
          inputMessage={inputMessage}
          onTabMove={() => {
            if (tabStart.current?.row !== row) tabStart.current = { row, col };
          }}
          onEnterAfterTabs={() => {
            const start = tabStart.current;
            tabStart.current = null;
            if (start?.row !== row) return false;
            goTo(single({ row: geo.rows.nextVisible(row + 1, 1) ?? row + 1, col: start.col }));
            return true;
          }}
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
            if (editing && pointMode) {
              // Typing a formula: keep editing (in the formula bar) and point at cells on that sheet.
              setEditing({ ...editing, source: "bar" });
              pendingCaret.current = editing.caret;
            } else if (editing) {
              commit(0, 0);
            }
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
        <StatusBar doc={doc} sheet={sheet} rect={dataRect()} mode={statusMode(editing, pointMode)} note={filterStatus} zoom={zoom} onZoom={setZoom} />
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
      {sizeDialog && (
        <SizeDialog
          axis={sizeDialog}
          px={sizeDialog === "rows" ? geo.rows.height(sel.top) : geo.cols.width(sel.left)}
          onApply={(px) => {
            const [first, last] = lineSpanOf(sizeDialog);
            if (sizeDialog === "rows") doc.setRowHeight(sheet, first, last, px);
            else doc.setColumnWidth(sheet, first, last, px);
            changed();
          }}
          onClose={() => {
            setSizeDialog(null);
            cellInputRef.current?.focus({ preventScroll: true });
          }}
        />
      )}
      {insertFunctionOpen && <InsertFunctionDialog onInsert={closeInsertFunction} onClose={() => closeInsertFunction(null)} />}
      {renderDataDialog()}
      {filterMenu && filter && (
        <FilterMenu
          key={`${sheet}:${filterMenu.col}`}
          doc={doc}
          sheet={sheet}
          filter={filter}
          col={filterMenu.col}
          x={filterMenu.x}
          y={filterMenu.y}
          onClose={() => {
            setFilterMenu(null);
            cellInputRef.current?.focus({ preventScroll: true });
          }}
          onChanged={changed}
          openDialog={setDataDialog}
        />
      )}
      {listMenu && listItems.length > 0 && (
        <ListDropdown
          items={listItems}
          x={listMenu.x}
          y={listMenu.y}
          minWidth={listMenu.width}
          onPick={pickFromList}
          onClose={() => {
            setListMenu(null);
            cellInputRef.current?.focus({ preventScroll: true });
          }}
        />
      )}
      {alertFor && <ValidationAlert rule={alertFor.rule} onChoice={answerAlert} />}
      {message && <MessageDialog title={message.title} message={message.message} onClose={() => setMessage(null)} />}
    </div>
  );
}

/** A toolbar colour swatch. Applies the colour once the picker closes. */
function ColorPicker({
  label,
  icon,
  title,
  value,
  onPick,
}: {
  label: string;
  /** Icon name from @office/ui, drawn above the colour bar instead of `label`. */
  icon?: string;
  title: string;
  value: string;
  onPick: (c: string) => void;
}) {
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
    <label className={icon ? "color-picker with-icon" : "color-picker"} title={icon ? undefined : title}>
      {icon ? (
        <span className="color-swatch">
          <Icon name={icon} />
          <span className="color-bar" style={{ background: value }} />
        </span>
      ) : (
        <span style={{ borderBottomColor: value }}>{label}</span>
      )}
      <input ref={ref} type="color" defaultValue={value} aria-label={title} />
    </label>
  );
}
