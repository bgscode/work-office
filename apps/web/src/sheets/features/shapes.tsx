// Insert → Shapes (a gallery of Excel's preset shapes, grouped like Excel's)
// and Text Box, and the Shape Format tab shown while a shape is selected:
// fill, outline, weight, change shape, text formatting and arrange. Shapes are
// drawn by dragging on the grid (see Grid.tsx); the engine keeps them
// (crates/sheets-engine/src/shape.rs).

import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { SheetsDocument } from "@office/engine";
import type { RibbonCommand, RibbonTab } from "@office/ui";
import { createPortal } from "react-dom";
import { drawShape } from "../charts/renderShape";
import { isLineShape, type ShapeInfo } from "../charts/types";
import { Select } from "../Select";

export interface ShapeContext {
  /** The selected shape, if a shape is selected. */
  shape: { id: number; info: ShapeInfo } | null;
  /** Start drawing `preset` ("textBox" for a text box) on the grid. */
  startDrawing: (preset: string) => void;
  openGallery: () => void;
  format: (key: string, value: string) => void;
  remove: () => void;
  reorder: (toFront: boolean) => void;
  editText: () => void;
}

/** Excel's theme colours (top row of its colour picker) and a few standard ones. */
const COLORS = [
  "#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
  "#C00000", "#FF0000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0", "#F2F2F2",
];

let presetCache: { group: string; preset: string; label: string }[] | null = null;
export function shapePresets() {
  if (!presetCache) {
    const flat = SheetsDocument.shapePresets();
    presetCache = [];
    for (let i = 0; i + 2 < flat.length; i += 3) presetCache.push({ group: flat[i], preset: flat[i + 1], label: flat[i + 2] });
  }
  return presetCache;
}

export function insertShapeCommands(ctx: ShapeContext): RibbonCommand[] {
  return [
    {
      id: "insert-shapes",
      label: "Shapes",
      title: "Insert ready-made shapes, such as circles, squares and arrows. Pick one, then drag on the sheet.",
      icon: "shapes",
      size: "large",
      group: "Illustrations",
      onRun: ctx.openGallery,
    },
    {
      id: "insert-textbox",
      label: "Text Box",
      title: "Draw a text box anywhere on the sheet.",
      icon: "textBox",
      size: "large",
      group: "Text",
      onRun: () => ctx.startDrawing("textBox"),
    },
  ];
}

function colorMenu(prefix: string, current: string | null, set: (c: string) => void, none: string | null) {
  return [
    ...(none ? [{ id: `${prefix}-none`, label: none, active: current === null, onRun: () => set("") }] : []),
    ...COLORS.map((c, i) => ({
      id: `${prefix}-${c}`,
      label: c,
      icon: <span className="menu-swatch" style={{ background: c }} />,
      active: current?.toUpperCase() === c,
      separatorBefore: i === 0 && !!none,
      onRun: () => set(c),
    })),
  ];
}

/** The contextual Shape Format tab, or null when no shape is selected. */
export function shapeFormatTab(ctx: ShapeContext): RibbonTab | null {
  const s = ctx.shape?.info;
  if (!s) return null;
  const line = isLineShape(s.preset);
  const family = shapePresets().filter((p) => isLineShape(p.preset) === line);
  const commands: RibbonCommand[] = [
    {
      id: "shape-change",
      label: "Edit Shape",
      title: "Change this shape to a different one, keeping its formatting and text.",
      icon: "shapes",
      size: "large",
      group: "Insert Shapes",
      menu: family.map((p, i) => ({
        id: `shape-to-${p.preset}`,
        label: p.label,
        active: p.preset === s.preset,
        separatorBefore: i > 0 && family[i - 1].group !== p.group,
        onRun: () => ctx.format("preset", p.preset),
      })),
    },
  ];
  if (!line) {
    commands.push({
      id: "shape-fill",
      label: "Shape Fill",
      title: "Fill the shape with a solid colour.",
      icon: "fillColor",
      size: "small",
      group: "Shape Styles",
      menu: colorMenu("fill", s.fill, (c) => ctx.format("fill", c), "No Fill"),
    });
  }
  commands.push(
    {
      id: "shape-outline",
      label: "Shape Outline",
      title: "Choose the colour of the shape's outline.",
      icon: "borders",
      size: "small",
      group: "Shape Styles",
      menu: colorMenu("line", s.line, (c) => ctx.format("line", c), line ? null : "No Outline"),
    },
    {
      id: "shape-weight",
      label: "Weight",
      title: "Choose the width of the outline.",
      icon: "lineWeight",
      size: "small",
      group: "Shape Styles",
      menu: [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => ({
        id: `weight-${w}`,
        label: `${w} pt`,
        active: Math.abs(s.lineWidth - w) < 0.01,
        onRun: () => ctx.format("lineWidth", String(w)),
      })),
    },
  );
  if (line) {
    commands.push({
      id: "shape-arrow",
      label: "Arrowhead",
      title: "Show an arrowhead at the end of the line.",
      icon: "arrowLine",
      size: "small",
      group: "Shape Styles",
      active: s.arrowEnd,
      onRun: () => ctx.format("arrowEnd", s.arrowEnd ? "0" : "1"),
    });
  } else {
    commands.push(
      { id: "shape-edit-text", label: "Edit Text", title: "Type text in the shape (or just start typing).", icon: "textBox", size: "large", group: "Text", onRun: ctx.editText },
      {
        id: "shape-font-size",
        label: "Font Size",
        group: "Text",
        render: () => (
          <Select aria-label="Shape font size" className="shape-font-size" value={String(s.fontSize)} onChange={(e) => ctx.format("fontSize", e.target.value)}>
            {[...new Set([8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72, s.fontSize])]
              .sort((a, b) => a - b)
              .map((n) => (
                <option key={n} value={String(n)}>
                  {n}
                </option>
              ))}
          </Select>
        ),
      },
      { id: "shape-bold", label: "Bold", icon: "bold", group: "Text", newRow: true, active: s.bold, onRun: () => ctx.format("bold", s.bold ? "0" : "1") },
      { id: "shape-italic", label: "Italic", icon: "italic", group: "Text", active: s.italic, onRun: () => ctx.format("italic", s.italic ? "0" : "1") },
      { id: "shape-left", label: "Align Left", icon: "alignLeft", group: "Text", active: s.align === "left", onRun: () => ctx.format("align", "left") },
      { id: "shape-center", label: "Center", icon: "alignCenter", group: "Text", active: s.align === "center", onRun: () => ctx.format("align", "center") },
      { id: "shape-right", label: "Align Right", icon: "alignRight", group: "Text", active: s.align === "right", onRun: () => ctx.format("align", "right") },
      {
        id: "shape-text-color",
        label: "Font Color",
        icon: "fontColor",
        group: "Text",
        menu: colorMenu("text", s.textColor, (c) => ctx.format("textColor", c), null),
      },
    );
  }
  commands.push(
    { id: "shape-front", label: "Bring to Front", icon: "bringFront", size: "small", group: "Arrange", onRun: () => ctx.reorder(true) },
    { id: "shape-back", label: "Send to Back", icon: "sendBack", size: "small", group: "Arrange", onRun: () => ctx.reorder(false) },
    { id: "shape-delete", label: "Delete", title: "Delete the shape (Delete).", icon: "trash", size: "small", group: "Arrange", onRun: ctx.remove },
  );
  return { id: "shape-format", label: "Shape Format", commands };
}

/** A shape drawn small, as an icon in the gallery and menus. */
export function ShapeIcon({ preset, size = 22 }: { preset: string; size?: number }) {
  const ref = useRef<HTMLCanvasElement>(null);
  useLayoutEffect(() => {
    const canvas = ref.current!;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = size * dpr;
    canvas.height = size * dpr;
    const ctx = canvas.getContext("2d")!;
    ctx.setTransform(dpr, 0, 0, dpr, 2, 2);
    const line = isLineShape(preset);
    drawShape(
      ctx,
      {
        preset,
        fill: null,
        line: "#444444",
        lineWidth: 0.75,
        arrowEnd: preset === "straightConnector1",
        flipV: line,
        text: "",
        textColor: "#000",
        fontSize: 11,
        bold: false,
        italic: false,
        align: "center",
        textTop: false,
      },
      size - 4 / dpr,
      preset.startsWith("wedge") || preset === "cloudCallout" ? size - 6 : size - 4 / dpr,
    );
  }, [preset, size]);
  return <canvas ref={ref} style={{ width: size, height: size }} aria-hidden />;
}

/** The Shapes gallery, under the Insert → Shapes button. */
export function ShapesGallery({ onPick, onClose }: { onPick: (preset: string) => void; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  useLayoutEffect(() => {
    const b = document.querySelector('[data-command="insert-shapes"]')?.getBoundingClientRect();
    setPos(b ? { left: Math.min(b.left, window.innerWidth - 330), top: b.bottom + 2 } : { left: 100, top: 120 });
  }, []);
  useEffect(() => {
    const onDown = (e: MouseEvent) => !ref.current?.contains(e.target as Node) && onClose();
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    const t = setTimeout(() => window.addEventListener("mousedown", onDown));
    window.addEventListener("keydown", onKey);
    return () => {
      clearTimeout(t);
      window.removeEventListener("mousedown", onDown);
      window.removeEventListener("keydown", onKey);
    };
  }, [onClose]);
  const groups: { name: string; items: { preset: string; label: string }[] }[] = [];
  for (const p of shapePresets()) {
    let g = groups.find((x) => x.name === p.group);
    if (!g) groups.push((g = { name: p.group, items: [] }));
    g.items.push(p);
  }
  if (!pos) return null;
  return createPortal(
    <div ref={ref} className="shapes-gallery" role="dialog" aria-label="Shapes" style={pos}>
      {groups.map((g) => (
        <section key={g.name}>
          <h4>{g.name}</h4>
          <div className="shapes-grid">
            {g.items.map((p) => (
              <button key={p.preset} title={p.label} aria-label={p.label} onClick={() => onPick(p.preset)}>
                <ShapeIcon preset={p.preset} />
              </button>
            ))}
          </div>
        </section>
      ))}
    </div>,
    document.body,
  );
}
