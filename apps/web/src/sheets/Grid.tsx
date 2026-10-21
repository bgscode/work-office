// The spreadsheet grid. Cells are painted on a <canvas> (only the visible
// ones, see paint.ts), and an invisible scrolling <div> on top provides native
// scrollbars and receives mouse events. Drawing cost depends on the window
// size, not on how much data there is.
//
// Zoom: everything is laid out in unzoomed "logical" pixels; only the canvas
// transform, the scroll spacer, mouse coordinates and the editor are scaled.
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
import type { SheetsDocument } from "@office/engine";
import {
  HEADER_HEIGHT,
  HEADER_WIDTH,
  MAX_COLS,
  MAX_ROWS,
  ROW_HEIGHT,
  clampCell,
  colName,
  moveSelection,
  selectionRect,
  single,
  type Cell,
  type Rect,
  type Selection,
} from "./geometry";
import { ColumnLayout } from "./layout";
import { CELL_FONT_SIZE, FILL_HANDLE_SIZE, colX, paint, rowY } from "./paint";
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
  | "pasteValuesNext";

export interface GridProps {
  doc: SheetsDocument;
  sheet: number;
  /** Bumped by the parent after every edit so the grid repaints. */
  revision: number;
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
  onContextMenu: (x: number, y: number) => void;
  onShortcut: (s: GridShortcut) => void;
  /** Ctrl+arrow: jump to the edge of the data (extend: also Shift). */
  onJump: (dRow: number, dCol: number, extend: boolean) => void;
}

type Hit =
  | { kind: "cell" | "col" | "row" | "all"; cell: Cell; onScrollbar: boolean }
  | { kind: "resize"; col: number; onScrollbar: false }
  | { kind: "fill"; onScrollbar: false };

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
  const [resizing, setResizing] = useState<{ col: number; width: number } | null>(null);
  const [fillPreview, setFillPreview] = useState<Rect | null>(null);
  const [cursor, setCursor] = useState("cell");
  const pendingReveal = useRef<Cell | null>(null);
  const editingInCell = editing?.source === "cell";

  const savedWidths = useMemo(() => doc.columnWidths(sheet), [doc, sheet, revision]);
  const layout = useMemo(() => {
    if (!resizing) return new ColumnLayout(savedWidths);
    return new ColumnLayout([...savedWidths, resizing.col, resizing.width]);
  }, [savedWidths, resizing]);

  const [dataRows, dataCols] = doc.extent(sheet);
  const rowCount = Math.min(MAX_ROWS, Math.max(1000, dataRows + 100, selection.anchor.row + 100, grow.rows));
  const colCount = Math.min(MAX_COLS, Math.max(30, dataCols + 10, selection.anchor.col + 10, grow.cols));
  const sel = selectionRect(selection);
  // Viewport size and scroll position in logical (unzoomed) pixels.
  const view = { width: size.width / zoom, height: size.height / zoom };
  const logicalScroll = { top: scroll.top / zoom, left: scroll.left / zoom };
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
      layout,
      highlights: props.highlights,
      fillPreview,
      showFillHandle,
    });
  }, [doc, sheet, revision, selection, scroll, size, zoom, layout, props.highlights, fillPreview, showFillHandle]);

  // Scroll a cell into view after the spacer has grown to include it.
  useLayoutEffect(() => {
    const cell = pendingReveal.current;
    const el = scrollerRef.current;
    if (!cell || !el) return;
    pendingReveal.current = null;
    const top = cell.row * ROW_HEIGHT;
    const left = layout.x(cell.col);
    const right = left + layout.width(cell.col);
    const viewH = el.clientHeight / zoom - HEADER_HEIGHT;
    const viewW = el.clientWidth / zoom - HEADER_WIDTH;
    const scrollTop = el.scrollTop / zoom;
    const scrollLeft = el.scrollLeft / zoom;
    if (top < scrollTop) el.scrollTop = top * zoom;
    else if (top + ROW_HEIGHT > scrollTop + viewH) el.scrollTop = (top + ROW_HEIGHT - viewH) * zoom;
    if (left < scrollLeft) el.scrollLeft = left * zoom;
    else if (right > scrollLeft + viewW) el.scrollLeft = (right - viewW) * zoom;
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

  function select(s: Selection, revealFocus = true) {
    props.onSelect(s);
    if (revealFocus) reveal(s.focus);
  }

  function onScroll() {
    const el = scrollerRef.current!;
    setScroll({ top: el.scrollTop, left: el.scrollLeft });
    const contentH = (HEADER_HEIGHT + rowCount * ROW_HEIGHT) * zoom;
    const contentW = (HEADER_WIDTH + layout.x(colCount)) * zoom;
    if (el.scrollTop + 2 * el.clientHeight > contentH) setGrow((g) => ({ ...g, rows: rowCount + 1000 }));
    if (el.scrollLeft + 2 * el.clientWidth > contentW) setGrow((g) => ({ ...g, cols: colCount + 20 }));
  }

  /** Logical viewport coordinates of a mouse event, relative to the grid's top-left. */
  function local(e: { clientX: number; clientY: number }) {
    const rect = scrollerRef.current!.getBoundingClientRect();
    return { x: (e.clientX - rect.left) / zoom, y: (e.clientY - rect.top) / zoom };
  }

  /** Current scroll position in logical pixels. */
  function scrolled() {
    const el = scrollerRef.current!;
    return { top: el.scrollTop / zoom, left: el.scrollLeft / zoom };
  }

  function cellAt(x: number, y: number): Cell {
    const sc = scrolled();
    return clampCell({
      row: Math.floor((y - HEADER_HEIGHT + sc.top) / ROW_HEIGHT),
      col: layout.colAt(x - HEADER_WIDTH + sc.left),
    });
  }

  function hitTest(e: { clientX: number; clientY: number }): Hit {
    const el = scrollerRef.current!;
    const { x, y } = local(e);
    const sc = scrolled();
    const onScrollbar = x > el.clientWidth / zoom || y > el.clientHeight / zoom;
    if (!onScrollbar && y < HEADER_HEIGHT && x >= HEADER_WIDTH) {
      // Near a column border in the header: resize.
      const contentX = x - HEADER_WIDTH + sc.left;
      const col = layout.colAt(contentX);
      if (Math.abs(contentX - layout.x(col + 1)) <= 4) return { kind: "resize", col, onScrollbar: false };
      if (col > 0 && Math.abs(contentX - layout.x(col)) <= 4) return { kind: "resize", col: col - 1, onScrollbar: false };
    }
    if (!onScrollbar && showFillHandle) {
      const hx = colX(layout, sel.right + 1, sc.left);
      const hy = rowY(sel.bottom + 1, sc.top);
      if (Math.abs(x - hx) <= FILL_HANDLE_SIZE && Math.abs(y - hy) <= FILL_HANDLE_SIZE) return { kind: "fill", onScrollbar: false };
    }
    const cell = cellAt(x, y);
    const kind = x < HEADER_WIDTH && y < HEADER_HEIGHT ? "all" : y < HEADER_HEIGHT ? "col" : x < HEADER_WIDTH ? "row" : "cell";
    return { kind, cell, onScrollbar };
  }

  /** Nudge the scroll position when dragging past the edge of the grid. */
  function autoScroll(e: globalThis.MouseEvent) {
    const el = scrollerRef.current!;
    const { x, y } = local(e);
    if (y > el.clientHeight / zoom - 8) el.scrollTop += ROW_HEIGHT;
    else if (y < HEADER_HEIGHT) el.scrollTop -= ROW_HEIGHT;
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
    if (e.button !== 0) return;
    const hit = hitTest(e);
    if (hit.onScrollbar) return;
    e.preventDefault();

    if (hit.kind === "resize") {
      const startX = e.clientX;
      const startWidth = layout.width(hit.col);
      let width = startWidth;
      setResizing({ col: hit.col, width });
      drag(
        (ev) => {
          width = Math.max(8, Math.round(startWidth + (ev.clientX - startX) / zoom));
          setResizing({ col: hit.col, width });
        },
        () => {
          setResizing(null);
          if (width !== startWidth) props.onResizeColumn(hit.col, width);
        },
      );
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
      props.onPoint(single(anchor), "start");
      drag((ev) => {
        const { x, y } = local(ev);
        props.onPoint({ anchor, focus: cellAt(x, y) }, "drag");
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
    props.onSelect(initial);

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
      props.onSelect({ anchor: initial.anchor, focus });
    });
  }

  function onHover(e: MouseEvent) {
    if (e.buttons) return;
    const hit = hitTest(e);
    const next = hit.kind === "resize" ? "col-resize" : hit.kind === "fill" ? "crosshair" : "cell";
    if (next !== cursor) setCursor(next);
  }

  function onDoubleClick(e: MouseEvent) {
    const hit = hitTest(e);
    if (hit.onScrollbar) return;
    if (hit.kind === "resize") props.onAutofit(hit.col);
    else if (hit.kind === "fill") props.onFillToEnd();
    else if (hit.kind === "cell") props.onStartEdit();
  }

  function onContextMenu(e: MouseEvent) {
    const hit = hitTest(e);
    if (hit.onScrollbar || hit.kind === "resize" || hit.kind === "fill") return;
    e.preventDefault();
    if (editing) props.onCommit(0, 0);
    // Right-clicking outside the selection selects what was clicked, like Excel.
    const c = hit.cell;
    const inside = c.row >= sel.top && c.row <= sel.bottom && c.col >= sel.left && c.col <= sel.right;
    if (!inside) {
      if (hit.kind === "col") props.onSelect({ anchor: { row: 0, col: c.col }, focus: { row: MAX_ROWS - 1, col: c.col } });
      else if (hit.kind === "row") props.onSelect({ anchor: { row: c.row, col: 0 }, focus: { row: c.row, col: MAX_COLS - 1 } });
      else props.onSelect(single(c));
    }
    props.onContextMenu(e.clientX, e.clientY);
  }

  function onKeyDown(e: KeyboardEvent<HTMLTextAreaElement>) {
    // Keys that confirm an IME composition belong to the IME, not to us.
    if (e.nativeEvent.isComposing) return;
    if (editingInCell) onEditKey(e);
    else onNavigateKey(e);
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
        reveal(moveSelection(selection, dr, dc, false).anchor);
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
        home: () => select(single({ row: 0, col: 0 })),
        end: () => {
          const [rows, cols] = doc.extent(sheet);
          select(single({ row: Math.max(0, rows - 1), col: Math.max(0, cols - 1) }));
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
      select(moveSelection(selection, dr, dc, e.shiftKey));
      return;
    }
    switch (e.key) {
      case "Enter":
        e.preventDefault();
        select(moveSelection(selection, e.shiftKey ? -1 : 1, 0, false));
        return;
      case "Tab":
        e.preventDefault();
        select(moveSelection(selection, 0, e.shiftKey ? -1 : 1, false));
        return;
      case "Home":
        e.preventDefault();
        select(single({ row: selection.anchor.row, col: 0 }));
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
    return Math.max(1, Math.floor((view.height - HEADER_HEIGHT) / ROW_HEIGHT) - 1);
  }

  const { row, col } = selection.anchor;
  const draft = editingInCell ? editing!.draft : "";
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
        <div style={{ width: (HEADER_WIDTH + layout.x(colCount)) * zoom, height: (HEADER_HEIGHT + rowCount * ROW_HEIGHT) * zoom }} />
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
          left: colX(layout, col, logicalScroll.left) * zoom,
          top: rowY(row, logicalScroll.top) * zoom,
          width: Math.max(layout.width(col) + 1, longestLine(draft) * 8 + 16) * zoom,
          height: (Math.max(1, draft.split("\n").length) * (ROW_HEIGHT - 4) + 5) * zoom,
          fontSize: CELL_FONT_SIZE * zoom,
          lineHeight: `${(ROW_HEIGHT - 4) * zoom}px`,
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
