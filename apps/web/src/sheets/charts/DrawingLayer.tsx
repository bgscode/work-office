// Charts and pictures floating over the grid. Each drawing is anchored to
// cells (see crates/sheets-engine/src/chart.rs), so it is positioned from the
// sheet's column/row layout and follows resizes, hidden lines and scrolling.
//
// Like Excel: click selects (frame with eight sizing handles), drag moves,
// handles resize (Shift keeps the shape), Delete removes, Esc goes back to the
// cells. A move or resize is one undoable step, sent when the mouse is released.

import { useEffect, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent as ReactKeyboardEvent } from "react";
import type { SheetsDocument } from "@office/engine";
import { HEADER_HEIGHT, HEADER_WIDTH } from "../geometry";
import type { SheetGeometry } from "../paint";
import { axisLengths, drawChart } from "./renderChart";
import { drawShape } from "./renderShape";
import { isLineShape, type AnchorArray, type ChartView, type DrawingInfo, type ShapeInfo } from "./types";

export interface Rect4 {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** A drawing's rectangle in sheet pixels (from the top-left of A1). */
export function anchorRect(geo: SheetGeometry, a: AnchorArray): Rect4 {
  const x = geo.cols.x(a[1]) + Math.min(a[3], geo.cols.width(a[1]));
  const y = geo.rows.y(a[0]) + Math.min(a[2], geo.rows.height(a[0]));
  const x2 = geo.cols.x(a[5]) + Math.min(a[7], geo.cols.width(a[5]));
  const y2 = geo.rows.y(a[4]) + Math.min(a[6], geo.rows.height(a[4]));
  return { x, y, w: Math.max(0, x2 - x), h: Math.max(0, y2 - y) };
}

/** The anchor for a rectangle in sheet pixels. */
export function rectAnchor(geo: SheetGeometry, r: Rect4): AnchorArray {
  const point = (x: number, y: number) => {
    const col = geo.cols.colAt(Math.max(0, x));
    const row = geo.rows.rowAt(Math.max(0, y));
    return [row, col, Math.max(0, y - geo.rows.y(row)), Math.max(0, x - geo.cols.x(col))];
  };
  return [...point(r.x, r.y), ...point(r.x + r.w, r.y + r.h)] as AnchorArray;
}

const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
type Handle = (typeof HANDLES)[number];
const MIN_SIZE = 20;

interface Props {
  doc: SheetsDocument;
  sheet: number;
  revision: number;
  geo: SheetGeometry;
  /** Scroll position in logical pixels. */
  scroll: { top: number; left: number };
  zoom: number;
  selected: number | null;
  onSelect: (id: number | null) => void;
  /** A drawing was moved or resized. */
  onAnchor: (id: number, anchor: AnchorArray) => void;
  onDelete: (id: number) => void;
  onContextMenu: (id: number, x: number, y: number) => void;
  /** Wheel over a drawing scrolls the grid. */
  onWheel: (dx: number, dy: number) => void;
  /** Keys a drawing doesn't handle itself (undo, redo). */
  onKey: (e: ReactKeyboardEvent) => void;
  /** New text typed into a shape or text box. */
  onShapeText: (id: number, text: string) => void;
  /** Start editing this shape's text (e.g. right after drawing a text box). */
  editText: { id: number; seq: number } | null;
  /** A chart title typed in place (double-click the title); "" returns to the automatic title. */
  onChartTitle: (id: number, title: string) => void;
}

/** Height of the title band at the top of a chart (see renderChart.ts). */
const TITLE_BAND = 40;

export function DrawingLayer(props: Props) {
  const { doc, sheet, revision, geo, scroll, zoom } = props;
  const drawings = useMemo(() => JSON.parse(doc.drawingsJson(sheet) || "[]") as DrawingInfo[], [doc, sheet, revision]);
  // Live rectangle while moving or resizing (sheet pixels).
  const [drag, setDrag] = useState<{ id: number; rect: Rect4 } | null>(null);
  const layerRef = useRef<HTMLDivElement>(null);
  // The shape whose text is being typed, and the text so far.
  const [textEdit, setTextEdit] = useState<{ id: number; text: string } | null>(null);
  useEffect(() => {
    if (!props.editText) return;
    const d = drawings.find((x) => x.id === props.editText!.id);
    if (d?.kind === "shape") setTextEdit({ id: d.id, text: d.shape.text });
  }, [props.editText?.seq]);
  // The chart whose title is being typed in place.
  const [titleEdit, setTitleEdit] = useState<{ id: number; text: string } | null>(null);
  function finishTitle() {
    if (!titleEdit) return;
    const d = drawings.find((x) => x.id === titleEdit.id);
    if (d?.kind === "chart" && titleEdit.text !== (d.chart.title ?? "")) props.onChartTitle(titleEdit.id, titleEdit.text);
    setTitleEdit(null);
  }

  function finishText() {
    if (!textEdit) return;
    const d = drawings.find((x) => x.id === textEdit.id);
    if (d?.kind === "shape" && d.shape.text !== textEdit.text) props.onShapeText(textEdit.id, textEdit.text);
    setTextEdit(null);
  }
  // A selected drawing has the keyboard (Delete, arrows, Esc), however it was selected.
  useEffect(() => {
    if (props.selected === null) return;
    layerRef.current?.querySelector<HTMLElement>(`[data-drawing="${props.selected}"]`)?.focus({ preventScroll: true });
  }, [props.selected, drawings]);
  if (drawings.length === 0) return null;

  /** Sheet pixels to layer pixels (the layer starts below/right of the headers). */
  const toLayer = (r: Rect4): Rect4 => {
    const f = geo.freeze;
    const frozenCols = f.cols > 0 && r.x < geo.cols.x(f.cols);
    const frozenRows = f.rows > 0 && r.y < geo.rows.y(f.rows);
    // Panes frozen while scrolled start at their first frozen line.
    const originX = f.cols > 0 ? geo.cols.x(f.left) : 0;
    const originY = f.rows > 0 ? geo.rows.y(f.top) : 0;
    return {
      x: (r.x - originX - (frozenCols ? 0 : scroll.left)) * zoom,
      y: (r.y - originY - (frozenRows ? 0 : scroll.top)) * zoom,
      w: r.w * zoom,
      h: r.h * zoom,
    };
  };

  function startDrag(e: React.MouseEvent, d: DrawingInfo, handle: Handle | null) {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    props.onSelect(d.id);
    (e.currentTarget as HTMLElement).closest<HTMLElement>(".drawing")?.focus({ preventScroll: true });
    const start = anchorRect(geo, d.anchor);
    // A horizontal or vertical line has no height or width.
    const min = d.kind === "shape" && isLineShape(d.shape.preset) ? 0 : MIN_SIZE;
    const ratio = start.w / Math.max(1, start.h);
    const x0 = e.clientX;
    const y0 = e.clientY;
    let rect = start;
    let moved = false;
    const move = (ev: MouseEvent) => {
      const dx = (ev.clientX - x0) / zoom;
      const dy = (ev.clientY - y0) / zoom;
      if (!moved && Math.abs(dx) < 3 && Math.abs(dy) < 3) return;
      moved = true;
      if (!handle) {
        rect = { ...start, x: Math.max(0, start.x + dx), y: Math.max(0, start.y + dy) };
      } else {
        let { x, y, w, h } = start;
        if (handle.includes("e")) w = Math.max(min, start.w + dx);
        if (handle.includes("s")) h = Math.max(min, start.h + dy);
        if (handle.includes("w")) {
          w = Math.max(min, start.w - dx);
          x = start.x + start.w - w;
        }
        if (handle.includes("n")) {
          h = Math.max(min, start.h - dy);
          y = start.y + start.h - h;
        }
        // Shift on a corner keeps the proportions.
        if (ev.shiftKey && handle.length === 2) {
          if (w / h > ratio) h = w / ratio;
          else w = h * ratio;
          if (handle.includes("n")) y = start.y + start.h - h;
          if (handle.includes("w")) x = start.x + start.w - w;
        }
        rect = { x: Math.max(0, x), y: Math.max(0, y), w, h };
      }
      setDrag({ id: d.id, rect });
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      setDrag(null);
      if (moved) props.onAnchor(d.id, rectAnchor(geo, rect));
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  function onKeyDown(e: ReactKeyboardEvent, d: DrawingInfo) {
    if (d.kind === "shape" && !isLineShape(d.shape.preset) && (e.key === "Enter" || e.key === "F2" || (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey))) {
      // Typing on a selected shape edits its text, like Excel.
      e.preventDefault();
      setTextEdit({ id: d.id, text: e.key.length === 1 ? d.shape.text + e.key : d.shape.text });
    } else if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      props.onDelete(d.id);
    } else if (e.key === "Escape") {
      e.preventDefault();
      props.onSelect(null);
    } else if (e.key.startsWith("Arrow")) {
      // Arrow keys nudge the drawing by a pixel (Excel moves by one cell edge with Alt; we keep it simple).
      e.preventDefault();
      const r = anchorRect(geo, d.anchor);
      const step = e.shiftKey ? 10 : 1;
      const dx = e.key === "ArrowLeft" ? -step : e.key === "ArrowRight" ? step : 0;
      const dy = e.key === "ArrowUp" ? -step : e.key === "ArrowDown" ? step : 0;
      props.onAnchor(d.id, rectAnchor(geo, { ...r, x: Math.max(0, r.x + dx), y: Math.max(0, r.y + dy) }));
    } else props.onKey(e);
  }

  return (
    <div ref={layerRef} className="drawing-layer" style={{ left: HEADER_WIDTH * zoom, top: HEADER_HEIGHT * zoom }} onWheel={(e) => props.onWheel(e.deltaX, e.deltaY)}>
      {drawings.map((d) => {
        const rect = drag?.id === d.id ? drag.rect : anchorRect(geo, d.anchor);
        const line = d.kind === "shape" && isLineShape(d.shape.preset);
        if (line ? rect.w <= 0 && rect.h <= 0 : rect.w <= 0 || rect.h <= 0) return null;
        // Lines get room around them to be seen and grabbed even when flat.
        const pad = line ? 6 : 0;
        const box = toLayer(rect);
        const px = { x: box.x - pad, y: box.y - pad, w: box.w + 2 * pad, h: box.h + 2 * pad };
        const selected = props.selected === d.id;
        return (
          <div
            key={d.id}
            data-drawing={d.id}
            className={selected ? "drawing selected" : "drawing"}
            role="img"
            aria-label={d.name}
            tabIndex={-1}
            style={{ left: px.x, top: px.y, width: px.w, height: px.h }}
            onMouseDown={(e) => startDrag(e, d, null)}
            onContextMenu={(e) => {
              e.preventDefault();
              props.onSelect(d.id);
              props.onContextMenu(d.id, e.clientX, e.clientY);
            }}
            onKeyDown={(e) => textEdit?.id !== d.id && onKeyDown(e, d)}
            onDoubleClick={(e) => {
              if (d.kind === "shape" && !isLineShape(d.shape.preset)) setTextEdit({ id: d.id, text: d.shape.text });
              else if (d.kind === "chart" && d.chart.showTitle) {
                // Double-clicking the title edits it in place, like Excel.
                const top = (e.currentTarget as HTMLElement).getBoundingClientRect().top;
                if ((e.clientY - top) / zoom > TITLE_BAND) return;
                const view = doc.chartViewJson(sheet, d.id, 100, 100);
                setTitleEdit({ id: d.id, text: (view && (JSON.parse(view) as ChartView).title) || "" });
              }
            }}
          >
            {d.kind === "chart" ? (
              <ChartCanvas doc={doc} sheet={sheet} id={d.id} revision={revision} width={rect.w} height={rect.h} zoom={zoom} />
            ) : d.kind === "shape" ? (
              <ShapeCanvas shape={textEdit?.id === d.id ? { ...d.shape, text: "" } : d.shape} width={rect.w} height={rect.h} zoom={zoom} pad={pad / zoom} />
            ) : (
              <Picture doc={doc} sheet={sheet} id={d.id} mime={d.mime} name={d.name} />
            )}
            {titleEdit?.id === d.id && (
              <input
                className="chart-title-editor"
                autoFocus
                aria-label="Chart title"
                value={titleEdit.text}
                style={{ fontSize: 18.5 * zoom, top: 8 * zoom, height: 30 * zoom }}
                onFocus={(e) => e.currentTarget.select()}
                onChange={(e) => setTitleEdit({ id: d.id, text: e.target.value })}
                onMouseDown={(e) => e.stopPropagation()}
                onKeyDown={(e) => {
                  e.stopPropagation();
                  if (e.key === "Enter" || e.key === "Escape") {
                    e.preventDefault();
                    if (e.key === "Escape") setTitleEdit(null);
                    else finishTitle();
                    layerRef.current?.querySelector<HTMLElement>(`[data-drawing="${d.id}"]`)?.focus();
                  }
                }}
                onBlur={finishTitle}
              />
            )}
            {textEdit?.id === d.id && d.kind === "shape" && (
              <textarea
                className="shape-text-editor"
                autoFocus
                aria-label="Shape text"
                value={textEdit.text}
                style={{
                  color: d.shape.textColor,
                  fontSize: d.shape.fontSize * (4 / 3) * zoom,
                  fontWeight: d.shape.bold ? 700 : 400,
                  fontStyle: d.shape.italic ? "italic" : "normal",
                  textAlign: d.shape.align,
                  paddingTop: d.shape.textTop ? 4.8 * zoom : Math.max(0, (rect.h * zoom) / 2 - d.shape.fontSize * 0.8 * zoom),
                }}
                onFocus={(e) => e.currentTarget.setSelectionRange(e.currentTarget.value.length, e.currentTarget.value.length)}
                onChange={(e) => setTextEdit({ id: d.id, text: e.target.value })}
                onMouseDown={(e) => e.stopPropagation()}
                onKeyDown={(e) => {
                  e.stopPropagation();
                  if (e.key === "Escape") {
                    e.preventDefault();
                    finishText();
                    layerRef.current?.querySelector<HTMLElement>(`[data-drawing="${d.id}"]`)?.focus();
                  }
                }}
                onBlur={finishText}
              />
            )}
            {selected &&
              textEdit?.id !== d.id &&
              (d.kind === "shape" && isLineShape(d.shape.preset) ? (d.shape.flipV ? (["sw", "ne"] as const) : (["nw", "se"] as const)) : HANDLES).map((h) => (
                <span
                  key={h}
                  className={`drawing-handle ${h}`}
                  style={line ? { left: pad + (h.includes("e") ? box.w : 0), top: pad + (h.includes("s") ? box.h : 0) } : undefined}
                  onMouseDown={(e) => startDrag(e, d, h)}
                />
              ))}
          </div>
        );
      })}
    </div>
  );
}

function ChartCanvas({ doc, sheet, id, revision, width, height, zoom }: { doc: SheetsDocument; sheet: number; id: number; revision: number; width: number; height: number; zoom: number }) {
  const ref = useRef<HTMLCanvasElement>(null);
  // Axis lengths only change the number of gridlines; round them so dragging doesn't refetch data constantly.
  const lengths = axisLengths(width, height);
  const valuePx = Math.round(lengths.value / 20) * 20;
  const xPx = Math.round(lengths.x / 20) * 20;
  const view = useMemo(() => {
    const json = doc.chartViewJson(sheet, id, valuePx, xPx);
    return json ? (JSON.parse(json) as ChartView) : null;
  }, [doc, sheet, id, revision, valuePx, xPx]);
  useLayoutEffect(() => {
    const canvas = ref.current;
    if (!canvas || !view) return;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.round(width * zoom * dpr));
    canvas.height = Math.max(1, Math.round(height * zoom * dpr));
    const ctx = canvas.getContext("2d")!;
    ctx.setTransform(dpr * zoom, 0, 0, dpr * zoom, 0, 0);
    drawChart(ctx, view, width, height);
  }, [view, width, height, zoom]);
  return <canvas ref={ref} className="drawing-canvas" />;
}

function ShapeCanvas({ shape, width, height, zoom, pad }: { shape: ShapeInfo; width: number; height: number; zoom: number; pad: number }) {
  const ref = useRef<HTMLCanvasElement>(null);
  useLayoutEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(1, Math.round((width + 2 * pad) * zoom * dpr));
    canvas.height = Math.max(1, Math.round((height + 2 * pad) * zoom * dpr));
    const ctx = canvas.getContext("2d")!;
    ctx.setTransform(dpr * zoom, 0, 0, dpr * zoom, pad * zoom * dpr, pad * zoom * dpr);
    drawShape(ctx, shape, width, height);
  }, [shape, width, height, zoom, pad]);
  return <canvas ref={ref} className="drawing-canvas" />;
}

function Picture({ doc, sheet, id, mime, name }: { doc: SheetsDocument; sheet: number; id: number; mime: string; name: string }) {
  const [url, setUrl] = useState<string | null>(null);
  useEffect(() => {
    const bytes = doc.pictureData(sheet, id);
    const u = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
    setUrl(u);
    return () => URL.revokeObjectURL(u);
  }, [doc, sheet, id, mime]);
  return url ? <img className="drawing-picture" src={url} alt={name} draggable={false} /> : null;
}
