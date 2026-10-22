// Home-tab formatting controls that go beyond on/off buttons: font and size
// boxes, Borders, vertical alignment, Wrap Text and Format Painter. SheetsApp
// owns the state and passes callbacks in; this file only builds ribbon commands.

import type { CellStyle } from "@office/engine";
import { Icon, type RibbonCommand, type RibbonMenuItem } from "@office/ui";
import { DEFAULT_FONT, DEFAULT_FONT_SIZE, FONT_CHOICES, FONT_SIZE_CHOICES, fontFamilyCss } from "../cellRenderer";
import { ComboBox } from "./ComboBox";

export type PainterMode = "off" | "once" | "sticky";

export interface BorderLineChoice {
  style: string;
  /** "#RRGGBB" or "" for automatic. */
  color: string;
}

export interface FormattingContext {
  style: CellStyle;
  /** `formatRange` key/value (see crates/wasm/src/lib.rs). */
  format: (key: string, value: string) => void;
  applyBorders: (preset: string) => void;
  borderLine: BorderLineChoice;
  setBorderLine: (line: BorderLineChoice) => void;
  /** The preset the Borders button repeats (Excel remembers the last one). */
  lastBorder: string;
  painter: PainterMode;
  setPainter: (mode: PainterMode) => void;
  focusGrid: () => void;
}

const BORDER_PRESETS: { id: string; label: string; separatorBefore?: boolean }[] = [
  { id: "bottom", label: "Bottom Border" },
  { id: "top", label: "Top Border" },
  { id: "left", label: "Left Border" },
  { id: "right", label: "Right Border" },
  { id: "none", label: "No Border", separatorBefore: true },
  { id: "all", label: "All Borders" },
  { id: "outside", label: "Outside Borders" },
  { id: "thickOutside", label: "Thick Outside Borders" },
  { id: "bottomDouble", label: "Bottom Double Border", separatorBefore: true },
  { id: "thickBottom", label: "Thick Bottom Border" },
  { id: "topBottom", label: "Top and Bottom Border" },
  { id: "topThickBottom", label: "Top and Thick Bottom Border" },
  { id: "topDoubleBottom", label: "Top and Double Bottom Border" },
];

const LINE_STYLES: { id: string; label: string }[] = [
  { id: "thin", label: "Thin line" },
  { id: "medium", label: "Medium line" },
  { id: "thick", label: "Thick line" },
  { id: "dashed", label: "Dashed line" },
  { id: "dotted", label: "Dotted line" },
  { id: "double", label: "Double line" },
];

const LINE_COLORS: { id: string; label: string }[] = [
  { id: "", label: "Automatic" },
  { id: "#C00000", label: "Dark Red" },
  { id: "#FF0000", label: "Red" },
  { id: "#FFC000", label: "Orange" },
  { id: "#00B050", label: "Green" },
  { id: "#0070C0", label: "Blue" },
  { id: "#7030A0", label: "Purple" },
  { id: "#808080", label: "Gray" },
];

/** Font box, size box, Increase/Decrease Font Size — the first row of Excel's Font group. */
export function fontCommands(ctx: FormattingContext): RibbonCommand[] {
  const font = ctx.style.font ?? DEFAULT_FONT;
  const size = String(ctx.style.sz ?? DEFAULT_FONT_SIZE);
  return [
    {
      id: "font-name",
      label: "Font",
      group: "Font",
      render: () => (
        <ComboBox
          label="Font"
          title="Font: pick a new font for your text."
          value={font}
          options={FONT_CHOICES}
          width={140}
          onCommit={(v) => ctx.format("font", v === DEFAULT_FONT ? "" : v)}
          onDone={ctx.focusGrid}
          renderOption={(o) => <span style={{ fontFamily: fontFamilyCss(o) }}>{o}</span>}
        />
      ),
    },
    {
      id: "font-size",
      label: "Font Size",
      group: "Font",
      render: () => (
        <ComboBox
          label="Font Size"
          title="Font Size: change the size of your text."
          value={size}
          options={FONT_SIZE_CHOICES.map(String)}
          width={52}
          onCommit={(v) => {
            const n = Number(v);
            if (n >= 1 && n <= 409) ctx.format("size", String(n));
          }}
          onDone={ctx.focusGrid}
        />
      ),
    },
    {
      id: "font-grow",
      label: "Increase Font Size",
      shortcut: "Ctrl+Shift+>",
      title: "Make your text a bit bigger.",
      icon: <span className="font-step">A<sup>▲</sup></span>,
      group: "Font",
      onRun: () => ctx.format("fontStep", "up"),
    },
    {
      id: "font-shrink",
      label: "Decrease Font Size",
      shortcut: "Ctrl+Shift+<",
      title: "Make your text a bit smaller.",
      icon: <span className="font-step small">A<sup>▼</sup></span>,
      group: "Font",
      onRun: () => ctx.format("fontStep", "down"),
    },
  ];
}

/** The Borders split button: click repeats the last preset, the arrow opens the menu. */
export function bordersCommand(ctx: FormattingContext): RibbonCommand {
  const lastLabel = BORDER_PRESETS.find((p) => p.id === ctx.lastBorder)?.label ?? "Bottom Border";
  const menu: RibbonMenuItem[] = [
    ...BORDER_PRESETS.map((p) => ({ id: `border-${p.id}`, label: p.label, separatorBefore: p.separatorBefore, onRun: () => ctx.applyBorders(p.id) })),
    ...LINE_STYLES.map((l, i) => ({
      id: `line-${l.id}`,
      label: l.label,
      active: ctx.borderLine.style === l.id,
      separatorBefore: i === 0,
      onRun: () => ctx.setBorderLine({ ...ctx.borderLine, style: l.id }),
    })),
    ...LINE_COLORS.map((c, i) => ({
      id: `color-${c.id || "auto"}`,
      label: `Line Color: ${c.label}`,
      active: ctx.borderLine.color === c.id,
      separatorBefore: i === 0,
      icon: <span className="color-chip" style={{ background: c.id || "#000" }} />,
      onRun: () => ctx.setBorderLine({ ...ctx.borderLine, color: c.id }),
    })),
  ];
  return {
    id: "borders",
    label: lastLabel,
    title: "Borders: apply borders to the selected cells.",
    icon: "borders",
    group: "Font",
    onRun: () => ctx.applyBorders(ctx.lastBorder),
    menu,
  };
}

/** Top/Middle/Bottom Align and Wrap Text — the first row of Excel's Alignment group. */
export function alignmentCommands(ctx: FormattingContext): RibbonCommand[] {
  const va = ctx.style.va ?? "bottom";
  return [
    { id: "valign-top", label: "Top Align", title: "Align text to the top.", icon: "alignTop", group: "Alignment", active: va === "top", onRun: () => ctx.format("valign", "top") },
    { id: "valign-middle", label: "Middle Align", title: "Align text so it's centered between the top and bottom of the cell.", icon: "alignMiddle", group: "Alignment", active: va === "center", onRun: () => ctx.format("valign", "center") },
    { id: "valign-bottom", label: "Bottom Align", title: "Align text to the bottom.", icon: "alignBottom", group: "Alignment", active: va === "bottom", onRun: () => ctx.format("valign", "bottom") },
    {
      id: "wrap",
      label: "Wrap Text",
      title: "Wrap extra-long text into multiple lines so you can see all of it.",
      icon: "wrap",
      size: "small",
      group: "Alignment",
      active: !!ctx.style.wrap,
      onRun: () => ctx.format("wrap", ctx.style.wrap ? "0" : "1"),
    },
  ];
}

/** Format Painter: click to paint once, double-click to keep painting until Esc. */
export function painterCommand(ctx: FormattingContext): RibbonCommand {
  return {
    id: "format-painter",
    label: "Format Painter",
    group: "Clipboard",
    render: () => (
      <button
        className={`ribbon-button painter${ctx.painter !== "off" ? " pressed" : ""}`}
        aria-pressed={ctx.painter !== "off"}
        title="Format Painter: copy formatting from one place and apply it to another. Double-click to apply it to several places; press Esc to stop."
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => ctx.setPainter(ctx.painter === "off" ? "once" : "off")}
        onDoubleClick={() => ctx.setPainter("sticky")}
      >
        <span className="ribbon-icon">
          <Icon name="formatPainter" size={16} />
        </span>
        <span className="ribbon-label">Format Painter</span>
      </button>
    ),
  };
}
