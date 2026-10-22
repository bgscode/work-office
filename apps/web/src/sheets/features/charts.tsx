// Ribbon commands and dialogs for charts and pictures: Insert → Charts and
// Illustrations, and the Chart Design tab shown while a chart is selected
// (Add Chart Element, Change Colors, Switch Row/Column, Select Data, Change
// Chart Type, Arrange). The engine does the work (crates/wasm/src/charts.rs).

import { useState } from "react";
import type { RibbonCommand, RibbonTab } from "@office/ui";
import { Dialog } from "../Dialog";
import { CHART_GROUPS, type ChartInfo, type ChartTypeId, type LegendPos } from "../charts/types";

export interface ChartContext {
  /** The selected chart, if a chart is selected. */
  chart: { id: number; info: ChartInfo } | null;
  insertChart: (type: ChartTypeId) => void;
  insertPicture: () => void;
  /** One `formatChart` change on the selected chart. */
  format: (key: string, value: string) => void;
  remove: () => void;
  reorder: (toFront: boolean) => void;
  openDialog: (d: ChartDialog) => void;
}

export type ChartDialog = "titles" | "selectData";

/** Excel's "Change Colors" palettes (Monochromatic palettes shade one accent). */
const COLOR_SETS: { id: string; label: string; colors: string[] | null }[] = [
  { id: "colorful", label: "Colorful (Office)", colors: null },
  { id: "blue", label: "Monochromatic Blue", colors: ["#264478", "#4472C4", "#8FAADC", "#B4C7E7", "#1F3864", "#2F5597"] },
  { id: "orange", label: "Monochromatic Orange", colors: ["#9E480E", "#ED7D31", "#F4B183", "#F8CBAD", "#843C0C", "#C55A11"] },
  { id: "green", label: "Monochromatic Green", colors: ["#43682B", "#70AD47", "#A9D18E", "#C5E0B4", "#385723", "#548235"] },
  { id: "gray", label: "Monochromatic Gray", colors: ["#3A3A3A", "#7F7F7F", "#A5A5A5", "#D9D9D9", "#262626", "#595959"] },
];

export function insertChartCommands(ctx: ChartContext): RibbonCommand[] {
  return [
    {
      id: "insert-picture",
      label: "Pictures",
      title: "Insert a picture from your computer (PNG, JPEG, GIF or BMP).",
      icon: "picture",
      size: "large",
      group: "Illustrations",
      onRun: ctx.insertPicture,
    },
    ...CHART_GROUPS.map(
      (g): RibbonCommand => ({
        id: `insert-${g.id}`,
        label: g.label,
        title: `Insert a ${g.label.toLowerCase()} chart from the selected data. A single cell uses the table around it.`,
        icon: g.icon,
        size: "large",
        group: "Charts",
        onRun: () => ctx.insertChart(g.types[0].id),
        menu: g.types.map((t) => ({ id: `insert-${t.id}`, label: t.label, onRun: () => ctx.insertChart(t.id) })),
      }),
    ),
  ];
}

/** The contextual Chart Design tab, or null when no chart is selected. */
export function chartDesignTab(ctx: ChartContext): RibbonTab | null {
  const c = ctx.chart?.info;
  if (!c) return null;
  const round = c.type === "pie" || c.type === "doughnut";
  const legend = (pos: LegendPos, label: string) => ({ id: `legend-${pos}`, label, active: c.legend === pos, onRun: () => ctx.format("legend", pos) });
  return {
    id: "chart-design",
    label: "Chart Design",
    commands: [
      {
        id: "chart-element",
        label: "Add Chart Element",
        title: "Add elements to the chart: title, axis titles, data labels, legend.",
        icon: "chartElement",
        size: "large",
        group: "Chart Layouts",
        menu: [
          { id: "el-title", label: "Chart Title", active: c.showTitle, onRun: () => ctx.format("showTitle", c.showTitle ? "0" : "1") },
          { id: "el-titles", label: "Edit Titles…", onRun: () => ctx.openDialog("titles") },
          { id: "el-xtitle", label: "Axis Title: Horizontal", active: c.xTitle !== null, disabled: round, onRun: () => ctx.format("xTitle", c.xTitle !== null ? "" : "Axis Title"), separatorBefore: true },
          { id: "el-ytitle", label: "Axis Title: Vertical", active: c.yTitle !== null, disabled: round, onRun: () => ctx.format("yTitle", c.yTitle !== null ? "" : "Axis Title") },
          { id: "el-labels", label: "Data Labels", active: c.dataLabels, onRun: () => ctx.format("dataLabels", c.dataLabels ? "0" : "1"), separatorBefore: true },
          { ...legend("none", "Legend: None"), separatorBefore: true },
          legend("right", "Legend: Right"),
          legend("top", "Legend: Top"),
          legend("left", "Legend: Left"),
          legend("bottom", "Legend: Bottom"),
        ],
      },
      {
        id: "chart-colors",
        label: "Change Colors",
        title: "Choose a different set of colours for the chart.",
        icon: "chartColors",
        size: "large",
        group: "Chart Styles",
        disabled: round,
        menu: COLOR_SETS.map((set) => ({
          id: `colors-${set.id}`,
          label: set.label,
          onRun: () => {
            const n = round ? 0 : c.series.length;
            for (let i = 0; i < Math.max(n, 1); i++) ctx.format("seriesColor", `${i}:${set.colors ? set.colors[i % set.colors.length] : ""}`);
          },
        })),
      },
      { id: "chart-switch", label: "Switch Row/Column", title: "Swap the data over the axes: rows become series and columns categories, or the reverse.", icon: "switchRowCol", size: "large", group: "Data", disabled: !c.source, onRun: () => ctx.format("switch", "") },
      { id: "chart-data", label: "Select Data", title: "Change the range of data the chart uses.", icon: "table", size: "large", group: "Data", onRun: () => ctx.openDialog("selectData") },
      {
        id: "chart-type",
        label: "Change Chart Type",
        title: "Change to a different type of chart.",
        icon: "chartColumn",
        size: "large",
        group: "Type",
        menu: CHART_GROUPS.flatMap((g, gi) =>
          g.types.map((t, ti) => ({ id: `type-${t.id}`, label: t.label, active: c.type === t.id, separatorBefore: gi > 0 && ti === 0, onRun: () => ctx.format("type", t.id) })),
        ),
      },
      { id: "chart-front", label: "Bring to Front", title: "Bring the chart in front of other objects.", icon: "bringFront", size: "small", group: "Arrange", onRun: () => ctx.reorder(true) },
      { id: "chart-back", label: "Send to Back", title: "Send the chart behind other objects.", icon: "sendBack", size: "small", group: "Arrange", onRun: () => ctx.reorder(false) },
      { id: "chart-delete", label: "Delete Chart", title: "Delete the selected chart (Delete).", icon: "trash", size: "small", group: "Arrange", onRun: ctx.remove },
    ],
  };
}

export function ChartTitlesDialog({ info, onApply, onClose }: { info: ChartInfo; onApply: (key: string, value: string) => void; onClose: () => void }) {
  const round = info.type === "pie" || info.type === "doughnut";
  const [title, setTitle] = useState(info.title ?? "");
  const [x, setX] = useState(info.xTitle ?? "");
  const [y, setY] = useState(info.yTitle ?? "");
  function ok() {
    if (title !== (info.title ?? "")) onApply("title", title);
    if (title && !info.showTitle) onApply("showTitle", "1");
    if (!round && x !== (info.xTitle ?? "")) onApply("xTitle", x);
    if (!round && y !== (info.yTitle ?? "")) onApply("yTitle", y);
    onClose();
  }
  const field = (label: string, value: string, set: (v: string) => void, hint: string, focus = false) => (
    <label className="dialog-field">
      <span>{label}</span>
      <input aria-label={label} autoFocus={focus} value={value} placeholder={hint} onChange={(e) => set(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} />
    </label>
  );
  return (
    <Dialog
      title="Chart Titles"
      onClose={onClose}
      width={380}
      footer={
        <>
          <button className="dialog-button primary" onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      {field("Chart title", title, setTitle, "Automatic", true)}
      {!round && field("Horizontal axis title", x, setX, "None")}
      {!round && field("Vertical axis title", y, setY, "None")}
    </Dialog>
  );
}

export function SelectDataDialog({ info, onApply, onClose }: { info: ChartInfo; onApply: (range: string) => boolean; onClose: () => void }) {
  const [range, setRange] = useState(info.source ? `=${info.source}` : "");
  const [error, setError] = useState("");
  function ok() {
    const r = range.trim().replace(/^=/, "");
    if (!r) return;
    if (!onApply(r)) {
      setError("That range has no numbers to chart. Use a range like A1:C10 or Sheet2!A1:C10.");
      return;
    }
    onClose();
  }
  return (
    <Dialog
      title="Select Data Source"
      onClose={onClose}
      width={460}
      footer={
        <>
          <button className="dialog-button primary" onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <label className="dialog-field">
        <span>Chart data range:</span>
        <input aria-label="Chart data range" autoFocus value={range} onChange={(e) => setRange(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} />
      </label>
      <p className="dialog-hint">Series: {info.series.map((s, i) => s.name?.replace(/^=/, "") || `Series${i + 1}`).join(", ") || "none"}</p>
      {error && <p className="dialog-error">{error}</p>}
    </Dialog>
  );
}
