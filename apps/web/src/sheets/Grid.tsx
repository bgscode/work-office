// The spreadsheet grid. Cells are painted on a <canvas> (only the visible
// ones, see paint.ts), and an invisible scrolling <div> on top provides native
// scrollbars and receives mouse events. Drawing cost depends on the window
// size, not on how much data there is.
//
// Geometry (row heights, column widths, hidden lines, merged cells, frozen
// panes) arrives as `geo` (see sheetGeometry.ts); `Panes` turns it into
// viewport positions. Everything is laid out in unzoomed "logical" pixels;
// only the canvas transform, the scroll spacer, mouse coordinates and the
// editor are scaled by `zoom`.
//
// Keyboard focus lives in a <textarea> that sits on the active cell. It is
// invisible until an edit starts, so typing (including IME input for Hindi,
// Chinese, etc.) goes straight into it, as in Google Sheets.

import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
  type MouseEvent,
  type RefObject,
} from "react";
import { parseStyle, type SheetsDocument } from "@office/engine";
import { editorFont } from "./cellRenderer";
import { HEADER_HEIGHT, HEADER_WIDTH, MAX_COLS, MAX_ROWS, colName, selectionRect, single, type Cell, type Rect, type Selection } from "./geometry";
import { DEFAULT_ROW_HEIGHT } from "./layout";
import { navigate } from "./merges";
import type { FilterInfo } from "./features/dataTools";
import { DrawingLayer } from "./charts/DrawingLayer";
import type { AnchorArray } from "./charts/types";
import { FILL_HANDLE_SIZE, filterButtonBox, makePanes, paint, type SheetGeometry } from "./paint";
import { Panes } from "./panes";
import { FORMULA_BAR_ID, type EditState } from "./types";

export type FillDirection = "down" | "up" | "right" | "left";
export type GridShortcut =
  | "bold"
  | "italic"
  | "underline"
  | "find"
  | "replace"
  | "goto"
  | "fillDown"
  | "fillRight"
  | "pasteValuesNext"
  | "filter"
  | "dropdown";

export interface GridProps {
  doc: SheetsDocument;
  sheet: number;
  /** Bumped by the parent after every edit so the grid repaints. */
  revision: number;
  /** Row/column sizes, hidden lines, merges and frozen panes. */
  geo: SheetGeometry;
  selection: Selection;
  editing: EditState | null;
  inputRef: RefObject<HTMLTextAreaElement>;
  /** 1 = 100%. */
  zoom: number;
  onZoom: (zoom: number) => void;
  /** Scroll a cell into view whenever `seq` changes (Go To, Find, Ctrl+arrow). */
  reveal: { row: number; col: number; seq: number } | null;
  /** While typing a formula at a spot where a reference fits, clicks insert references. */
  pointMode: boolean;
  /** References in the formula being edited. */
  highlights: Rect[];
  onSelect: (s: Selection) => void;
  /** `initial` replaces the cell content (typing); omitted means edit the existing input. */
  onStartEdit: (initial?: string) => void;
  onDraftChange: (draft: string, caret: number) => void;
  onCaret: (caret: number) => void;
  /** Gets first look at keys while editing (autocomplete); returns true if it handled the key. */
  onEditorKey: (e: KeyboardEvent<HTMLTextAreaElement>) => boolean;
  onCommit: (dRow: number, dCol: number) => void;
  onCancel: () => void;
  onPoint: (s: Selection, phase: "start" | "drag") => void;
  onPointMove: (dRow: number, dCol: number, extend: boolean) => void;
  onClear: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onCopy: () => string;
  onCut: () => string;
  onPaste: (text: string) => void;
  onFill: (direction: FillDirection, count: number, increment: boolean) => void;
  onFillToEnd: () => void;
  onResizeColumn: (col: number, px: number) => void;
  onAutofit: (col: number) => void;
  onResizeRow: (row: number, px: number) => void;
  onAutofitRow: (row: number) => void;
  onContextMenu: (x: number, y: number) => void;
  onShortcut: (s: GridShortcut) => void;
  /** Ctrl+arrow: jump to the edge of the data (extend: also Shift). */
  onJump: (dRow: number, dCol: number, extend: boolean) => void;
  /** AutoFilter, for the header dropdown buttons. */
  filter: FilterInfo | null;
  /** A filter button was clicked; x/y is where its menu opens (screen coordinates). */
  onFilterButton: (col: number, x: number, y: number) => void;
  /** Circle Invalid Data. */
  circles: Cell[];
  /** The active cell has an in-cell list (Data Validation): show its arrow button. */
  listArrow: boolean;
  /** The list arrow was clicked; x/y is where the list opens, width the cell's. */
  onListArrow: (x: number, y: number, width: number) => void;
  /** Data Validation input message for the active cell. */
  inputMessage: { title: string; message: string } | null;
  /** Tab moved right (not editing): remember where a run of Tab entries started. */
  onTabMove: () => void;
  /** Enter (not editing): returns true when it went back to where the Tabs started. */
  onEnterAfterTabs: () => boolean;
  /** Charts and pictures: the selected one, and what happens to them. */
  selectedDrawing: number | null;
  onSelectDrawing: (id: number | null) => void;
  onDrawingAnchor: (id: number, anchor: AnchorArray) => void;
  onDeleteDrawing: (id: number) => void;
  onDrawingMenu: (id: number, x: number, y: number) => void;
}

type Hit =
  | { kind: "cell" | "col" | "row" | "all"; cell: Cell; onScrollbar: boolean }
  | { kind: "resize"; col: number; onScrollbar: false }
  | { kind: "rowResize"; row: number; onScrollbar: false }
  | { kind: "fill"; onScrollbar: false }
  | { kind: "filter"; col: number; x: number; y: number; onScrollbar: false };

const ARROWS: Record<string, [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

export function Grid(props: GridProps) {
  const { doc, sheet, revision, selection, editing, inputRef, zoom } = props;
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const scrollerRef = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ width: 0, height: 0 });
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  // The scrollable area grows as the user scrolls or navigates, up to Excel's limits.
  const [grow, setGrow] = useState({ rows: 0, cols: 0 });
  const [resizing, setResizing] = useState<{ axis: "col" | "row"; index: number; size: number } | null>(null);
  const [fillPreview, setFillPreview] = useState<Rect | null>(null);
  const [cursor, setCursor] = useState("cell");
  const pendingReveal = useRef<Cell | null>(null);
  const editingInCell = editing?.source === "cell";

  // While dragging a header border, preview the new size.
  const geo: SheetGeometry = useMemo(() => {
    if (!resizing) return props.geo;
    return resizing.axis === "col"
      ? { ...props.geo, cols: props.geo.cols.withSize(resizing.index, resizing.size) }
      : { ...props.geo, rows: props.geo.rows.withSize(resizing.index, resizing.size) };
  }, [props.geo, resizing]);
  const { rows, cols, merges } = geo;

  const [dataRows, dataCols] = doc.extent(sheet);
  const rowCount = Math.min(MAX_ROWS, Math.max(1000, dataRows + 100, selection.anchor.row + 100, grow.rows));
  const colCount = Math.min(MAX_COLS, Math.max(30, dataCols + 10, selection.anchor.col + 10, grow.cols));
  const sel = selectionRect(selection);
  // Viewport size and scroll position in logical (unzoomed) pixels.
  const view = { width: size.width / zoom, height: size.height / zoom };
  const logicalScroll = { top: scroll.top / zoom, left: scroll.left / zoom };
  const panes = makePanes(geo, logicalScroll, view);
  const showFillHandle = sel.bottom - sel.top < MAX_ROWS - 1 && sel.right - sel.left < MAX_COLS - 1 && !editing;

  useLayoutEffect(() => {
    const el = scrollerRef.current!;
    const observer = new ResizeObserver(() => setSize({ width: el.clientWidth, height: el.clientHeight }));
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  useLayoutEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || size.width === 0) return;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = size.width * dpr;
    canvas.height = size.height * dpr;
    canvas.style.width = `${size.width}px`;
    canvas.style.height = `${size.height}px`;
    const ctx = canvas.getContext("2d")!;
    ctx.setTransform(dpr * zoom, 0, 0, dpr * zoom, 0, 0);
    paint(ctx, {
      width: view.width,
      height: view.height,
      scroll: logicalScroll,
      doc,
      sheet,
      selection,
      geo,
      highlights: props.highlights,
      fillPreview,
      showFillHandle,
      filter: props.filter,
      circles: props.circles,
    });
  }, [doc, sheet, revision, selection, scroll, size, zoom, geo, props.highlights, fillPreview, showFillHandle, props.filter, props.circles]);

  // Scroll a cell (or its merged area) into view after the spacer has grown to include it.
  useLayoutEffect(() => {
    const cell = pendingReveal.current;
    const el = scrollerRef.current;
    if (!cell || !el) return;
    pendingReveal.current = null;
    const area = merges.cellRect(cell);
    const p = makePanes(geo, { top: el.scrollTop / zoom, left: el.scrollLeft / zoom }, view);
    const viewH = el.clientHeight / zoom - HEADER_HEIGHT;
    const viewW = el.clientWidth / zoom - HEADER_WIDTH;
    const top = Panes.revealOffset(rows, area.top, geo.freeze.rows, p.frozenHeight, el.scrollTop / zoom, viewH);
    const left = Panes.revealOffset(cols, area.left, geo.freeze.cols, p.frozenWidth, el.scrollLeft / zoom, viewW);
    if (top !== null) el.scrollTop = top * zoom;
    if (left !== null) el.scrollLeft = left * zoom;
  });

  useLayoutEffect(() => {
    if (props.reveal) reveal({ row: props.reveal.row, col: props.reveal.col });
  }, [props.reveal?.seq]);

  // Ctrl/⌘ + mouse wheel zooms the grid instead of the page.
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  useEffect(() => {
    const el = scrollerRef.current!;
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const next = Math.round(Math.min(2, Math.max(0.5, zoomRef.current * (e.deltaY < 0 ? 1.1 : 1 / 1.1))) * 100) / 100;
      latest.current.onZoom(next);
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, []);

  // Focus the cell input whenever we are not editing in the formula bar, and
  // put the caret at the end when an edit of existing content starts.
  useLayoutEffect(() => {
    const el = inputRef.current;
    if (!el || editing?.source === "bar") return;
    el.focus({ preventScroll: true });
    if (editingInCell) el.setSelectionRange(el.value.length, el.value.length);
  }, [editing === null, editing?.source, editing?.mode]);

  // Clipboard events go to the document; handle them while the grid has focus
  // and no text edit is in progress (then the input handles them natively).
  const latest = useRef(props);
  latest.current = props;
  useEffect(() => {
    const active = () => document.activeElement === inputRef.current && !latest.current.editing;
    const onCopy = (e: ClipboardEvent) => {
      if (!active()) return;
      e.preventDefault();
      e.clipboardData?.setData("text/plain", latest.current.onCopy());
    };
    const onCut = (e: ClipboardEvent) => {
      if (!active()) return;
      e.preventDefault();
      e.clipboardData?.setData("text/plain", latest.current.onCut());
    };
    const onPaste = (e: ClipboardEvent) => {
      if (!active()) return;
      e.preventDefault();
      latest.current.onPaste(e.clipboardData?.getData("text/plain") ?? "");
    };
    document.addEventListener("copy", onCopy);
    document.addEventListener("cut", onCut);
    document.addEventListener("paste", onPaste);
    return () => {
      document.removeEventListener("copy", onCopy);
      document.removeEventListener("cut", onCut);
      document.removeEventListener("paste", onPaste);
    };
  }, []);

  function reveal(cell: Cell) {
    pendingReveal.current = cell;
    setGrow((g) => ({ rows: Math.max(g.rows, cell.row + 100), cols: Math.max(g.cols, cell.col + 10) }));
  }

  /** Select, snapping to merged cells like Excel. */
  function select(s: Selection, revealFocus = true) {
    props.onSelect(merges.normalize(s));
    if (revealFocus) reveal(s.focus);
  }

  function onScroll() {
    const el = scrollerRef.current!;
    setScroll({ top: el.scrollTop, left: el.scrollLeft });
    const contentH = (HEADER_HEIGHT + rows.y(rowCount)) * zoom;
    const contentW = (HEADER_WIDTH + cols.x(colCount)) * zoom;
    if (el.scrollTop + 2 * el.clientHeight > contentH) setGrow((g) => ({ ...g, rows: rowCount + 1000 }));
    if (el.scrollLeft + 2 * el.clientWidth > contentW) setGrow((g) => ({ ...g, cols: colCount + 20 }));
  }

  /** Logical viewport coordinates of a mouse event, relative to the grid's top-left. */
  function local(e: { clientX: number; clientY: number }) {
    const rect = scrollerRef.current!.getBoundingClientRect();
    return { x: (e.clientX - rect.left) / zoom, y: (e.clientY - rect.top) / zoom };
  }

  /** Panes for the scroll position right now (the state may lag during a drag). */
  function livePanes(): Panes {
    const el = scrollerRef.current!;
    return makePanes(geo, { top: el.scrollTop / zoom, left: el.scrollLeft / zoom }, view);
  }

  function cellAt(x: number, y: number): Cell {
    const p = livePanes();
    return { row: Math.min(Math.max(p.rowAt(y), 0), MAX_ROWS - 1), col: Math.min(Math.max(p.colAt(x), 0), MAX_COLS - 1) };
  }

  function hitTest(e: { clientX: number; clientY: number }): Hit {
    const el = scrollerRef.current!;
    const { x, y } = local(e);
    const p = livePanes();
    const onScrollbar = x > el.clientWidth / zoom || y > el.clientHeight / zoom;
    if (!onScrollbar && y < HEADER_HEIGHT && x >= HEADER_WIDTH) {
      // Near a column border in the header: resize the column to its left.
      const col = p.colAt(x);
      const right = p.colX(col) + cols.width(col);
      if (Math.abs(x - right) <= 4) return { kind: "resize", col, onScrollbar: false };
      const prev = col > 0 ? cols.nextVisible(col - 1, -1) : null;
      if (prev !== null && Math.abs(x - p.colX(col)) <= 4) return { kind: "resize", col: prev, onScrollbar: false };
    }
    if (!onScrollbar && x < HEADER_WIDTH && y >= HEADER_HEIGHT) {
      // Near a row border in the header: resize the row above it.
      const row = p.rowAt(y);
      const bottom = p.rowY(row) + rows.height(row);
      if (Math.abs(y - bottom) <= 3) return { kind: "rowResize", row, onScrollbar: false };
      const prev = row > 0 ? rows.nextVisible(row - 1, -1) : null;
      if (prev !== null && Math.abs(y - p.rowY(row)) <= 3) return { kind: "rowResize", row: prev, onScrollbar: false };
    }
    if (!onScrollbar && showFillHandle) {
      const hx = p.colX(sel.right) + cols.width(sel.right);
      const hy = p.rowY(sel.bottom) + rows.height(sel.bottom);
      if (Math.abs(x - hx) <= FILL_HANDLE_SIZE && Math.abs(y - hy) <= FILL_HANDLE_SIZE) return { kind: "fill", onScrollbar: false };
    }
    const cell = cellAt(x, y);
    const button = onScrollbar ? null : filterButtonAt(cell, x, y, p);
    if (button) return button;
    const kind = x < HEADER_WIDTH && y < HEADER_HEIGHT ? "all" : y < HEADER_HEIGHT ? "col" : x < HEADER_WIDTH ? "row" : "cell";
    return { kind, cell, onScrollbar };
  }

  /** The AutoFilter button under (x, y) in `cell`, if any; mirrors paint.ts. */
  function filterButtonAt(cell: Cell, x: number, y: number, p: Panes): Hit | null {
    const f = props.filter;
    if (!f || x < HEADER_WIDTH || y < HEADER_HEIGHT) return null;
    const area = merges.cellRect(cell);
    if (area.top > f.range.top || area.bottom < f.range.top || cell.col < f.range.left || cell.col > f.range.right) return null;
    const col = area.left === area.right && area.top === area.bottom ? cell.col : Math.min(area.right, f.range.right);
    const left = area.left === area.right ? cell.col : area.left;
    const box = filterButtonBox({
      x: p.colX(left),
      y: p.rowY(area.top),
      w: cols.x(Math.min(area.right, f.range.right) + 1) - cols.x(left),
      h: rows.y(area.bottom + 1) - rows.y(area.top),
    });
    if (x < box.x || x > box.x + box.size || y < box.y || y > box.y + box.size) return null;
    const r = scrollerRef.current!.getBoundingClientRect();
    return { kind: "filter", col, x: r.left + box.x * zoom, y: r.top + (box.y + box.size) * zoom + 1, onScrollbar: false };
  }

  /** Nudge the scroll position when dragging past the edge of the grid. */
  function autoScroll(e: globalThis.MouseEvent) {
    const el = scrollerRef.current!;
    const { x, y } = local(e);
    if (y > el.clientHeight / zoom - 8) el.scrollTop += DEFAULT_ROW_HEIGHT;
    else if (y < HEADER_HEIGHT) el.scrollTop -= DEFAULT_ROW_HEIGHT;
    if (x > el.clientWidth / zoom - 8) el.scrollLeft += 40;
    else if (x < HEADER_WIDTH) el.scrollLeft -= 40;
  }

  function drag(onMove: (e: globalThis.MouseEvent) => void, onUp?: (e: globalThis.MouseEvent) => void) {
    const move = (e: globalThis.MouseEvent) => {
      autoScroll(e);
      onMove(e);
    };
    const up = (e: globalThis.MouseEvent) => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      onUp?.(e);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  function onMouseDown(e: MouseEvent) {
    if (props.selectedDrawing !== null) props.onSelectDrawing(null);
    if (e.button !== 0) return;
    const hit = hitTest(e);
    if (hit.onScrollbar) return;
    e.preventDefault();

    if (hit.kind === "resize" || hit.kind === "rowResize") {
      const isCol = hit.kind === "resize";
      const index = isCol ? hit.col : hit.row;
      const start = isCol ? e.clientX : e.clientY;
      const startSize = isCol ? cols.width(index) : rows.height(index);
      let next = startSize;
      setResizing({ axis: isCol ? "col" : "row", index, size: next });
      drag(
        (ev) => {
          next = Math.max(isCol ? 8 : 4, Math.round(startSize + ((isCol ? ev.clientX : ev.clientY) - start) / zoom));
          setResizing({ axis: isCol ? "col" : "row", index, size: next });
        },
        () => {
          setResizing(null);
          if (next === startSize) return;
          if (isCol) props.onResizeColumn(index, next);
          else props.onResizeRow(index, next);
        },
      );
      return;
    }

    if (hit.kind === "filter") {
      if (editing) props.onCommit(0, 0);
      props.onFilterButton(hit.col, hit.x, hit.y);
      return;
    }

    if (hit.kind === "fill") {
      let target: { direction: FillDirection; count: number } | null = null;
      drag(
        (ev) => {
          const { x, y } = local(ev);
          const c = cellAt(x, y);
          const down = c.row - sel.bottom;
          const up = sel.top - c.row;
          const right = c.col - sel.right;
          const left = sel.left - c.col;
          const best = Math.max(down, up, right, left);
          if (best <= 0) target = null;
          else if (best === down) target = { direction: "down", count: down };
          else if (best === up) target = { direction: "up", count: up };
          else if (best === right) target = { direction: "right", count: right };
          else target = { direction: "left", count: left };
          setFillPreview(target && fillArea(sel, target.direction, target.count));
        },
        (ev) => {
          setFillPreview(null);
          if (target) props.onFill(target.direction, target.count, ev.ctrlKey || ev.altKey || ev.metaKey);
        },
      );
      return;
    }

    // While typing a formula, clicking cells inserts references instead of moving.
    if (props.pointMode && hit.kind === "cell") {
      const anchor = hit.cell;
      props.onPoint(merges.normalize(single(anchor)), "start");
      drag((ev) => {
        const { x, y } = local(ev);
        props.onPoint(merges.normalize({ anchor, focus: cellAt(x, y) }), "drag");
      });
      return;
    }

    if (editing) props.onCommit(0, 0);
    inputRef.current?.focus({ preventScroll: true });
    const span = (c: Cell): Selection => {
      switch (hit.kind) {
        case "all":
          return { anchor: { row: 0, col: 0 }, focus: { row: MAX_ROWS - 1, col: MAX_COLS - 1 } };
        case "col":
          return { anchor: { row: 0, col: c.col }, focus: { row: MAX_ROWS - 1, col: c.col } };
        case "row":
          return { anchor: { row: c.row, col: 0 }, focus: { row: c.row, col: MAX_COLS - 1 } };
        default:
          return single(c);
      }
    };
    const initial = e.shiftKey && hit.kind === "cell" ? { anchor: selection.anchor, focus: hit.cell } : span(hit.cell);
    select(initial, false);

    // Dragging moves the focus corner; row/column drags keep their full span.
    drag((ev) => {
      const { x, y } = local(ev);
      const c = cellAt(x, y);
      const focus =
        hit.kind === "col"
          ? { row: MAX_ROWS - 1, col: c.col }
          : hit.kind === "row"
            ? { row: c.row, col: MAX_COLS - 1 }
            : hit.kind === "all"
              ? initial.focus
              : c;
      select({ anchor: initial.anchor, focus }, false);
    });
  }

  function onHover(e: MouseEvent) {
    if (e.buttons) return;
    const hit = hitTest(e);
    const next =
      hit.kind === "resize" ? "col-resize" : hit.kind === "rowResize" ? "row-resize" : hit.kind === "fill" ? "crosshair" : hit.kind === "filter" ? "default" : "cell";
    if (next !== cursor) setCursor(next);
  }

  function onDoubleClick(e: MouseEvent) {
    const hit = hitTest(e);
    if (hit.onScrollbar || hit.kind === "filter") return;
    if (hit.kind === "resize") props.onAutofit(hit.col);
    else if (hit.kind === "rowResize") props.onAutofitRow(hit.row);
    else if (hit.kind === "fill") props.onFillToEnd();
    else if (hit.kind === "cell") props.onStartEdit();
  }

  function onContextMenu(e: MouseEvent) {
    const hit = hitTest(e);
    if (hit.onScrollbar || hit.kind === "resize" || hit.kind === "rowResize" || hit.kind === "fill" || hit.kind === "filter") return;
    e.preventDefault();
    if (editing) props.onCommit(0, 0);
    // Right-clicking outside the selection selects what was clicked, like Excel.
    const c = hit.cell;
    const inside = c.row >= sel.top && c.row <= sel.bottom && c.col >= sel.left && c.col <= sel.right;
    if (!inside) {
      if (hit.kind === "col") select({ anchor: { row: 0, col: c.col }, focus: { row: MAX_ROWS - 1, col: c.col } }, false);
      else if (hit.kind === "row") select({ anchor: { row: c.row, col: 0 }, focus: { row: c.row, col: MAX_COLS - 1 } }, false);
      else select(single(c), false);
    }
    props.onContextMenu(e.clientX, e.clientY);
  }

  function onKeyDown(e: KeyboardEvent<HTMLTextAreaElement>) {
    // Keys that confirm an IME composition belong to the IME, not to us.
    if (e.nativeEvent.isComposing) return;
    if (editingInCell) onEditKey(e);
    else onNavigateKey(e);
  }

  /** Arrow-key style move that skips hidden lines and treats merged cells as one cell. */
  function move(dRow: number, dCol: number, extend: boolean): Selection {
    return navigate(selection, dRow, dCol, extend, { rows, cols, merges });
  }

  function onEditKey(e: KeyboardEvent<HTMLTextAreaElement>) {
    if (props.onEditorKey(e)) return;
    if (e.key === "Enter") {
      e.preventDefault();
      props.onCommit(e.shiftKey ? -1 : 1, 0);
    } else if (e.key === "Tab") {
      e.preventDefault();
      props.onCommit(0, e.shiftKey ? -1 : 1);
    } else if (e.key === "Escape") {
      e.preventDefault();
      props.onCancel();
    } else if (editing?.mode === "enter" && e.key in ARROWS) {
      e.preventDefault();
      const [dr, dc] = ARROWS[e.key];
      if (props.pointMode) {
        props.onPointMove(dr, dc, e.shiftKey);
      } else {
        props.onCommit(dr, dc);
        reveal(move(dr, dc, false).anchor);
      }
    }
  }

  function onNavigateKey(e: KeyboardEvent) {
    const mod = e.metaKey || e.ctrlKey;
    const key = e.key.toLowerCase();
    if (mod && e.shiftKey && key === "v") {
      // Ctrl+Shift+V: let the browser fire "paste", but paste values only.
      props.onShortcut("pasteValuesNext");
      return;
    }
    if (mod && e.shiftKey && key === "l") {
      // Ctrl+Shift+L: Data → Filter.
      e.preventDefault();
      props.onShortcut("filter");
      return;
    }
    if (e.altKey && e.key === "ArrowDown") {
      // Alt+Down: open the cell's list or the column's filter dropdown.
      e.preventDefault();
      props.onShortcut("dropdown");
      return;
    }
    if (mod && e.key in ARROWS) {
      e.preventDefault();
      const [dr, dc] = ARROWS[e.key];
      props.onJump(dr, dc, e.shiftKey);
      return;
    }
    if (e.shiftKey && !mod && e.key === " ") {
      e.preventDefault();
      select({ anchor: { row: selection.anchor.row, col: 0 }, focus: { row: selection.focus.row, col: MAX_COLS - 1 } }, false);
      return;
    }
    if (e.key === "F5") {
      e.preventDefault();
      props.onShortcut("goto");
      return;
    }
    if (mod) {
      const shortcuts: Record<string, () => void> = {
        home: () => select(single({ row: geo.freeze.rows, col: geo.freeze.cols })),
        end: () => {
          const [extentRows, extentCols] = doc.extent(sheet);
          select(single({ row: Math.max(0, extentRows - 1), col: Math.max(0, extentCols - 1) }));
        },
        " ": () =>
          select({ anchor: { row: 0, col: selection.anchor.col }, focus: { row: MAX_ROWS - 1, col: selection.focus.col } }, false),
        g: () => props.onShortcut("goto"),
        h: () => props.onShortcut("replace"),
        z: () => (e.shiftKey ? props.onRedo() : props.onUndo()),
        y: props.onRedo,
        a: () => select({ anchor: { row: 0, col: 0 }, focus: { row: MAX_ROWS - 1, col: MAX_COLS - 1 } }, false),
        b: () => props.onShortcut("bold"),
        i: () => props.onShortcut("italic"),
        u: () => props.onShortcut("underline"),
        f: () => props.onShortcut("find"),
        d: () => props.onShortcut("fillDown"),
        r: () => props.onShortcut("fillRight"),
      };
      if (key in shortcuts) {
        e.preventDefault();
        shortcuts[key]();
      }
      return; // let copy/cut/paste and other browser shortcuts through
    }
    if (e.altKey) return;

    const moves: Record<string, [number, number]> = {
      ...ARROWS,
      PageUp: [-pageRows(), 0],
      PageDown: [pageRows(), 0],
    };
    if (e.key in moves) {
      e.preventDefault();
      const [dr, dc] = moves[e.key];
      const next = move(dr, dc, e.shiftKey);
      props.onSelect(next);
      reveal(next.focus);
      return;
    }
    switch (e.key) {
      case "Enter": {
        e.preventDefault();
        if (!e.shiftKey && props.onEnterAfterTabs()) return;
        const next = move(e.shiftKey ? -1 : 1, 0, false);
        props.onSelect(next);
        reveal(next.focus);
        return;
      }
      case "Tab": {
        e.preventDefault();
        if (!e.shiftKey) props.onTabMove();
        const next = move(0, e.shiftKey ? -1 : 1, false);
        props.onSelect(next);
        reveal(next.focus);
        return;
      }
      case "Home":
        e.preventDefault();
        select(single({ row: selection.anchor.row, col: cols.nextVisible(0, 1) ?? 0 }));
        return;
      case "F2":
        e.preventDefault();
        props.onStartEdit();
        return;
      case "Delete":
        e.preventDefault();
        props.onClear();
        return;
      case "Backspace":
        e.preventDefault();
        props.onStartEdit("");
        return;
    }
    // Printable keys fall through to the input; onChange starts the edit.
  }

  function pageRows() {
    return Math.max(1, Math.floor((view.height - HEADER_HEIGHT - panes.frozenHeight) / DEFAULT_ROW_HEIGHT) - 1);
  }

  // The editor covers the active cell, or its whole merged area.
  const { row, col } = selection.anchor;
  const area = merges.cellRect(selection.anchor);
  const draft = editingInCell ? editing!.draft : "";
  const cellStyle = parseStyle(doc.styleBlock(sheet, row, col, 1, 1)[0]);
  const font = editorFont(cellStyle);
  const areaW = cols.x(area.right + 1) - cols.x(area.left);
  const areaH = rows.y(area.bottom + 1) - rows.y(area.top);
  const lines = Math.max(1, draft.split("\n").length);
  const cellLeft = panes.colX(area.left) * zoom;
  const cellTop = panes.rowY(area.top) * zoom;
  return (
    <div className="grid">
      <canvas ref={canvasRef} className="grid-canvas" />
      <div
        ref={scrollerRef}
        className="grid-scroller"
        role="grid"
        aria-label="Spreadsheet"
        style={{ cursor: props.pointMode && cursor === "cell" ? "copy" : cursor }}
        onScroll={onScroll}
        onMouseDown={onMouseDown}
        onMouseMove={onHover}
        onDoubleClick={onDoubleClick}
        onContextMenu={onContextMenu}
      >
        <div style={{ width: (HEADER_WIDTH + cols.x(colCount)) * zoom, height: (HEADER_HEIGHT + rows.y(rowCount)) * zoom }} />
      </div>
      <div className="drawing-clip" style={{ width: Math.max(0, size.width), height: Math.max(0, size.height) }}>
        <DrawingLayer
          doc={doc}
          sheet={sheet}
          revision={revision}
          geo={geo}
          scroll={logicalScroll}
          zoom={zoom}
          selected={props.selectedDrawing}
          onSelect={(id) => {
            props.onSelectDrawing(id);
            if (id === null) inputRef.current?.focus({ preventScroll: true });
          }}
          onAnchor={props.onDrawingAnchor}
          onDelete={(id) => {
            props.onDeleteDrawing(id);
            inputRef.current?.focus({ preventScroll: true });
          }}
          onContextMenu={props.onDrawingMenu}
          onWheel={(dx, dy) => scrollerRef.current?.scrollBy(dx, dy)}
          onKey={(e) => {
            const mod = e.metaKey || e.ctrlKey;
            const key = e.key.toLowerCase();
            if (mod && (key === "z" || key === "y")) {
              e.preventDefault();
              if (key === "y" || e.shiftKey) props.onRedo();
              else props.onUndo();
            }
          }}
        />
      </div>
      <textarea
        ref={inputRef}
        className={editingInCell ? "cell-editor" : "cell-editor idle"}
        aria-label={`Cell ${colName(col)}${row + 1}`}
        spellCheck={false}
        autoComplete="off"
        rows={1}
        value={draft}
        style={{
          left: panes.colX(area.left) * zoom,
          top: panes.rowY(area.top) * zoom,
          width: Math.max(areaW + 1, longestLine(draft) * font.fontSize * 0.55 + 16) * zoom,
          height: Math.max(areaH + 1, lines * font.lineHeight + 4) * zoom,
          fontFamily: font.fontFamily,
          fontSize: font.fontSize * zoom,
          fontWeight: font.fontWeight,
          fontStyle: font.fontStyle,
          lineHeight: `${font.lineHeight * zoom}px`,
        }}
        onChange={(e) =>
          editing ? props.onDraftChange(e.target.value, e.target.selectionStart ?? e.target.value.length) : props.onStartEdit(e.target.value)
        }
        onSelect={(e) => editingInCell && props.onCaret(e.currentTarget.selectionStart ?? 0)}
        onKeyDown={onKeyDown}
        onBlur={(e) => {
          if (editingInCell && e.relatedTarget?.id !== FORMULA_BAR_ID) props.onCommit(0, 0);
        }}
      />
      {props.listArrow && !editingInCell && (
        <button
          className="list-arrow"
          aria-label="Show list"
          tabIndex={-1}
          style={{ left: cellLeft + areaW * zoom + 1, top: cellTop, height: Math.min(areaH, 20) * zoom }}
          onMouseDown={(e) => {
            e.preventDefault();
            const r = scrollerRef.current!.getBoundingClientRect();
            props.onListArrow(r.left + cellLeft, r.top + cellTop + areaH * zoom + 1, areaW * zoom);
          }}
        >
          <svg width="8" height="5" viewBox="0 0 8 5" aria-hidden>
            <path d="M0 0h8L4 5z" fill="currentColor" />
          </svg>
        </button>
      )}
      {props.inputMessage && !editing && (
        <div className="input-message" role="tooltip" style={{ left: cellLeft + (areaW * zoom) / 2, top: cellTop + areaH * zoom + 6 }}>
          {props.inputMessage.title && <strong>{props.inputMessage.title}</strong>}
          {props.inputMessage.message}
        </div>
      )}
    </div>
  );
}

function longestLine(text: string): number {
  return Math.max(...text.split("\n").map((l) => l.length));
}

/** The cells a fill-handle drag covers beyond the selection. */
export function fillArea(sel: Rect, direction: FillDirection, count: number): Rect {
  switch (direction) {
    case "down":
      return { ...sel, top: sel.bottom + 1, bottom: Math.min(MAX_ROWS - 1, sel.bottom + count) };
    case "up":
      return { ...sel, top: Math.max(0, sel.top - count), bottom: sel.top - 1 };
    case "right":
      return { ...sel, left: sel.right + 1, right: Math.min(MAX_COLS - 1, sel.right + count) };
    case "left":
      return { ...sel, left: Math.max(0, sel.left - count), right: sel.left - 1 };
  }
}
