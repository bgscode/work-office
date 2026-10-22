// JSON shapes written by crates/wasm/src/charts.rs.

/** `[fromRow, fromCol, fromRowOffPx, fromColOffPx, toRow, toCol, toRowOffPx, toColOffPx]` */
export type AnchorArray = [number, number, number, number, number, number, number, number];

export type ChartTypeId =
  | "column"
  | "column-stacked"
  | "column-percent"
  | "bar"
  | "bar-stacked"
  | "bar-percent"
  | "line"
  | "line-markers"
  | "pie"
  | "doughnut"
  | "area"
  | "area-stacked"
  | "scatter"
  | "scatter-lines";

export type LegendPos = "none" | "right" | "bottom" | "top" | "left";

export interface ChartInfo {
  type: ChartTypeId;
  showTitle: boolean;
  title: string | null;
  xTitle: string | null;
  yTitle: string | null;
  legend: LegendPos;
  dataLabels: boolean;
  /** The data range, e.g. "A1:C5" or "Sheet2!A1:C5". */
  source: string | null;
  byRows: boolean;
  series: { name: string | null; values: string; categories: string | null; color: string | null }[];
}

/** A shape or text box (see crates/sheets-engine/src/shape.rs). */
export interface ShapeInfo {
  /** DrawingML preset geometry, e.g. "rect", "ellipse", "rightArrow", "line". */
  preset: string;
  fill: string | null;
  line: string | null;
  /** Points. */
  lineWidth: number;
  arrowEnd: boolean;
  flipV: boolean;
  text: string;
  textColor: string;
  /** Points. */
  fontSize: number;
  bold: boolean;
  italic: boolean;
  align: "left" | "center" | "right";
  textTop: boolean;
}

export type DrawingInfo =
  | { id: number; name: string; anchor: AnchorArray; kind: "chart"; chart: ChartInfo }
  | { id: number; name: string; anchor: AnchorArray; kind: "picture"; mime: string; bytes: number }
  | { id: number; name: string; anchor: AnchorArray; kind: "shape"; shape: ShapeInfo };

export const isLineShape = (preset: string) => preset === "line" || preset === "straightConnector1";

export interface AxisView {
  min: number;
  max: number;
  step: number;
  labels: string[];
}

export interface SeriesView {
  name: string;
  color: string;
  values: (number | null)[];
  labels: string[];
  xs: (number | null)[];
}

export interface ChartView {
  type: ChartTypeId;
  title: string | null;
  xTitle: string | null;
  yTitle: string | null;
  legend: LegendPos;
  dataLabels: boolean;
  categories: string[];
  series: SeriesView[];
  pointColors: string[];
  valueAxis: AxisView | null;
  xAxis: AxisView | null;
}

/** Chart types offered on Insert → Charts, grouped like Excel's gallery buttons. */
export const CHART_GROUPS: { id: string; label: string; icon: string; types: { id: ChartTypeId; label: string }[] }[] = [
  {
    id: "column",
    label: "Column",
    icon: "chartColumn",
    types: [
      { id: "column", label: "Clustered Column" },
      { id: "column-stacked", label: "Stacked Column" },
      { id: "column-percent", label: "100% Stacked Column" },
    ],
  },
  {
    id: "bar",
    label: "Bar",
    icon: "chartBar",
    types: [
      { id: "bar", label: "Clustered Bar" },
      { id: "bar-stacked", label: "Stacked Bar" },
      { id: "bar-percent", label: "100% Stacked Bar" },
    ],
  },
  {
    id: "line",
    label: "Line",
    icon: "chartLine",
    types: [
      { id: "line", label: "Line" },
      { id: "line-markers", label: "Line with Markers" },
      { id: "area", label: "Area" },
      { id: "area-stacked", label: "Stacked Area" },
    ],
  },
  {
    id: "pie",
    label: "Pie",
    icon: "chartPie",
    types: [
      { id: "pie", label: "Pie" },
      { id: "doughnut", label: "Doughnut" },
    ],
  },
  {
    id: "scatter",
    label: "Scatter",
    icon: "chartScatter",
    types: [
      { id: "scatter", label: "Scatter" },
      { id: "scatter-lines", label: "Scatter with Straight Lines" },
    ],
  },
];

export const CHART_TYPE_LABEL: Record<ChartTypeId, string> = Object.fromEntries(
  CHART_GROUPS.flatMap((g) => g.types.map((t) => [t.id, t.label])),
) as Record<ChartTypeId, string>;
