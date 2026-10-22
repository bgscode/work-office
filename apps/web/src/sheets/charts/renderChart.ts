// Draws a chart (ChartView from the engine) on a canvas, in Excel's default
// style: white chart area with a light grey border, grey title and axis text,
// light gridlines, the Office palette, Excel's gap widths and overlaps.
// Pure function of its inputs; DrawingLayer.tsx calls it whenever the chart,
// its data or its size changes.
//
// Big data: when there are more points than pixels, columns and lines are
// reduced to one min/max pair per pixel column, so a 100,000-row series draws
// as fast as a 500-row one and looks the same.

import type { AxisView, ChartView, LegendPos, SeriesView } from "./types";

const FONT = "Calibri, Carlito, 'Segoe UI', Arial, sans-serif";
const TEXT = "#595959";
const GRID = "#d9d9d9";
const BORDER = "#d9d9d9";

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Approximate axis lengths for `chartViewJson`, which decide how many gridlines fit. */
export function axisLengths(w: number, h: number): { value: number; x: number } {
  return { value: Math.max(40, h - 100), x: Math.max(40, w - 90) };
}

const isBar = (t: string) => t.startsWith("bar");
const isColumn = (t: string) => t.startsWith("column");
const isLine = (t: string) => t.startsWith("line");
const isArea = (t: string) => t.startsWith("area");
const isRound = (t: string) => t === "pie" || t === "doughnut";
const isScatter = (t: string) => t.startsWith("scatter");
const grouping = (t: string) => (t.endsWith("-percent") ? "percent" : t.endsWith("-stacked") ? "stacked" : "standard");

export function drawChart(ctx: CanvasRenderingContext2D, view: ChartView, w: number, h: number) {
  ctx.save();
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, w, h);
  ctx.strokeStyle = BORDER;
  ctx.lineWidth = 1;
  ctx.strokeRect(0.5, 0.5, w - 1, h - 1);
  ctx.textBaseline = "middle";

  let area: Box = { x: 10, y: 10, w: w - 20, h: h - 20 };
  if (view.title) {
    ctx.font = `18.5px ${FONT}`;
    ctx.fillStyle = TEXT;
    ctx.textAlign = "center";
    ctx.fillText(ellipsize(ctx, view.title, w - 20), w / 2, 22);
    area = { ...area, y: 40, h: area.h - 30 };
  }
  area = drawLegend(ctx, view, area);
  if (area.w < 30 || area.h < 30) {
    ctx.restore();
    return;
  }
  if (isRound(view.type)) drawPie(ctx, view, area);
  else drawAxesChart(ctx, view, area);
  ctx.restore();
}

function ellipsize(ctx: CanvasRenderingContext2D, text: string, max: number): string {
  if (ctx.measureText(text).width <= max) return text;
  let t = text;
  while (t.length > 1 && ctx.measureText(t + "…").width > max) t = t.slice(0, -1);
  return t + "…";
}

// ---------------------------------------------------------------- legend

interface LegendEntry {
  label: string;
  color: string;
  line: boolean;
}

function legendEntries(view: ChartView): LegendEntry[] {
  if (isRound(view.type)) return view.categories.map((c, i) => ({ label: c, color: view.pointColors[i], line: false }));
  const line = isLine(view.type) || view.type === "scatter-lines";
  // Excel lists bar-chart series bottom-up to match the bars' order.
  const list = view.series.map((s) => ({ label: s.name, color: s.color, line }));
  return isBar(view.type) ? list.reverse() : list;
}

function drawLegend(ctx: CanvasRenderingContext2D, view: ChartView, area: Box): Box {
  const pos: LegendPos = view.legend;
  if (pos === "none") return area;
  const entries = legendEntries(view).slice(0, 50);
  if (!entries.length) return area;
  ctx.font = `12px ${FONT}`;
  ctx.textAlign = "left";
  const sw = 8;
  const widths = entries.map((e) => sw + 5 + Math.min(ctx.measureText(e.label).width, 140) + 14);
  const drawEntry = (e: LegendEntry, x: number, y: number, maxW: number) => {
    ctx.fillStyle = e.color;
    ctx.strokeStyle = e.color;
    if (e.line) {
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + 14, y);
      ctx.stroke();
      x += 6;
    } else ctx.fillRect(x, y - sw / 2, sw, sw);
    ctx.fillStyle = TEXT;
    ctx.fillText(ellipsize(ctx, e.label, maxW), x + sw + 5, y);
  };
  if (pos === "bottom" || pos === "top") {
    // One or more centred rows.
    const rows: LegendEntry[][] = [[]];
    const rowWidths = [0];
    entries.forEach((e, i) => {
      if (rowWidths[rowWidths.length - 1] + widths[i] > area.w && rows[rows.length - 1].length) {
        rows.push([]);
        rowWidths.push(0);
      }
      rows[rows.length - 1].push(e);
      rowWidths[rowWidths.length - 1] += widths[i];
    });
    const shown = rows.slice(0, 3);
    const height = shown.length * 18 + 6;
    const top = pos === "bottom" ? area.y + area.h - height + 6 : area.y;
    let k = 0;
    shown.forEach((row, r) => {
      let x = area.x + (area.w - rowWidths[r]) / 2;
      for (const e of row) {
        drawEntry(e, x, top + r * 18 + 6, 140);
        x += widths[k++];
      }
    });
    return pos === "bottom" ? { ...area, h: area.h - height } : { ...area, y: area.y + height, h: area.h - height };
  }
  const colW = Math.min(Math.max(...widths), area.w / 3);
  const height = entries.length * 18;
  let y = area.y + Math.max(0, (area.h - height) / 2) + 9;
  const x = pos === "right" ? area.x + area.w - colW + 4 : area.x;
  for (const e of entries) {
    if (y > area.y + area.h) break;
    drawEntry(e, x, y, colW - sw - 14);
    y += 18;
  }
  return pos === "right" ? { ...area, w: area.w - colW - 4 } : { ...area, x: area.x + colW + 4, w: area.w - colW - 4 };
}

// ---------------------------------------------------------------- pie and doughnut

function drawPie(ctx: CanvasRenderingContext2D, view: ChartView, area: Box) {
  const s = view.series[0];
  if (!s) return;
  const values = s.values.map((v) => (v !== null && v > 0 ? v : 0));
  const total = values.reduce((a, b) => a + b, 0);
  if (total <= 0) return;
  const r = Math.max(10, Math.min(area.w, area.h) / 2 - 6);
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  const hole = view.type === "doughnut" ? r * 0.5 : 0;
  let angle = -Math.PI / 2;
  const mids: { a: number; i: number }[] = [];
  values.forEach((v, i) => {
    if (v <= 0) return;
    const sweep = (v / total) * Math.PI * 2;
    ctx.beginPath();
    if (hole) {
      ctx.arc(cx, cy, r, angle, angle + sweep);
      ctx.arc(cx, cy, hole, angle + sweep, angle, true);
    } else {
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, angle, angle + sweep);
    }
    ctx.closePath();
    ctx.fillStyle = view.pointColors[i] ?? s.color;
    ctx.fill();
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 1;
    ctx.stroke();
    mids.push({ a: angle + sweep / 2, i });
    angle += sweep;
  });
  if (view.dataLabels) {
    ctx.font = `12px ${FONT}`;
    ctx.textAlign = "center";
    ctx.fillStyle = hole ? TEXT : "#ffffff";
    const lr = hole ? (r + hole) / 2 : r * 0.65;
    for (const { a, i } of mids) ctx.fillText(s.labels[i] || String(s.values[i]), cx + Math.cos(a) * lr, cy + Math.sin(a) * lr);
  }
}

// ---------------------------------------------------------------- charts with axes

/** Values per series after stacking: [from, to] for each point, or null. */
function stacked(view: ChartView): ([number, number] | null)[][] {
  const g = grouping(view.type);
  const n = Math.max(0, ...view.series.map((s) => s.values.length));
  if (g === "standard") return view.series.map((s) => s.values.map((v) => (v === null ? null : [0, v])));
  const totals = Array.from({ length: n }, (_, i) => view.series.reduce((sum, s) => sum + Math.abs(s.values[i] ?? 0), 0));
  const pos = new Array(n).fill(0);
  const neg = new Array(n).fill(0);
  return view.series.map((s) =>
    s.values.map((v, i) => {
      if (v === null) return null;
      const x = g === "percent" ? (totals[i] ? v / totals[i] : 0) : v;
      const base = x >= 0 ? pos : neg;
      const from = base[i];
      base[i] += x;
      return [from, base[i]];
    }),
  );
}

function drawAxesChart(ctx: CanvasRenderingContext2D, view: ChartView, area: Box) {
  const axis = view.valueAxis;
  if (!axis) return;
  const horizontal = isBar(view.type);
  ctx.font = `12px ${FONT}`;
  const valueLabelW = Math.max(...axis.labels.map((l) => ctx.measureText(l).width), 10);
  const cats = view.categories;
  const xAxis = view.xAxis;

  // Room for axis titles and labels.
  let plot: Box = { ...area };
  if (view.yTitle) plot = { ...plot, x: plot.x + 20, w: plot.w - 20 };
  if (view.xTitle) plot = { ...plot, h: plot.h - 20 };
  const catLabelW = horizontal ? Math.min(120, Math.max(...cats.slice(0, 500).map((c) => ctx.measureText(c).width), 10)) : 0;
  const leftPad = (horizontal ? catLabelW : valueLabelW) + 8;
  plot = { x: plot.x + leftPad, y: plot.y + 6, w: plot.w - leftPad - 8, h: plot.h - 6 - 22 };
  if (plot.w < 20 || plot.h < 20) return;

  // Value scale: vertical (columns, lines, areas, scatter y) or horizontal (bars).
  const vlen = horizontal ? plot.w : plot.h;
  const vpos = (v: number) => ((v - axis.min) / (axis.max - axis.min || 1)) * vlen;
  const valueAt = horizontal ? (v: number) => plot.x + vpos(v) : (v: number) => plot.y + plot.h - vpos(v);

  // Gridlines and value labels.
  ctx.strokeStyle = GRID;
  ctx.lineWidth = 1;
  ctx.fillStyle = TEXT;
  axis.labels.forEach((label, i) => {
    const v = axis.min + i * axis.step;
    const p = Math.round(valueAt(v)) + 0.5;
    ctx.beginPath();
    if (horizontal) {
      ctx.moveTo(p, plot.y);
      ctx.lineTo(p, plot.y + plot.h);
    } else {
      ctx.moveTo(plot.x, p);
      ctx.lineTo(plot.x + plot.w, p);
    }
    ctx.stroke();
    if (horizontal) {
      ctx.textAlign = "center";
      ctx.fillText(label, p, plot.y + plot.h + 12);
    } else {
      ctx.textAlign = "right";
      ctx.fillText(label, plot.x - 6, p);
    }
  });

  if (isScatter(view.type) && xAxis) {
    drawScatter(ctx, view, plot, axis, xAxis, valueAt);
  } else {
    drawCategoryChart(ctx, view, plot, horizontal, valueAt);
  }

  // Axis titles.
  ctx.fillStyle = TEXT;
  ctx.font = `13px ${FONT}`;
  ctx.textAlign = "center";
  if (view.xTitle) ctx.fillText(view.xTitle, plot.x + plot.w / 2, area.y + area.h - 8);
  if (view.yTitle) {
    ctx.save();
    ctx.translate(area.x + 8, plot.y + plot.h / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText(view.yTitle, 0, 0);
    ctx.restore();
  }
}

function drawCategoryChart(ctx: CanvasRenderingContext2D, view: ChartView, plot: Box, horizontal: boolean, valueAt: (v: number) => number) {
  const n = Math.max(1, view.categories.length);
  const clen = horizontal ? plot.h : plot.w;
  const step = clen / n;
  // Bars run top-down in category order reversed, like Excel (first category at the bottom).
  const catStart = (i: number) => (horizontal ? plot.y + plot.h - (i + 1) * step : plot.x + i * step);
  const catMid = (i: number) => catStart(i) + step / 2;
  const segs = stacked(view);
  const g = grouping(view.type);

  // Category axis line and labels (thinned to what fits).
  ctx.strokeStyle = GRID;
  ctx.beginPath();
  const base = Math.round(valueAt(Math.max(view.valueAxis!.min, Math.min(0, view.valueAxis!.max)))) + 0.5;
  if (horizontal) {
    ctx.moveTo(base, plot.y);
    ctx.lineTo(base, plot.y + plot.h);
  } else {
    ctx.moveTo(plot.x, base);
    ctx.lineTo(plot.x + plot.w, base);
  }
  ctx.stroke();
  ctx.font = `12px ${FONT}`;
  ctx.fillStyle = TEXT;
  const labelSpace = horizontal ? 16 : Math.max(24, Math.min(90, ...view.categories.slice(0, 200).map((c) => ctx.measureText(c).width + 8)));
  const every = Math.max(1, Math.ceil(labelSpace / step));
  for (let i = 0; i < view.categories.length; i += every) {
    if (horizontal) {
      ctx.textAlign = "right";
      ctx.fillText(ellipsize(ctx, view.categories[i], 118), plot.x - 6, catMid(i));
    } else {
      ctx.textAlign = "center";
      ctx.fillText(ellipsize(ctx, view.categories[i], step * every - 4), catMid(i), plot.y + plot.h + 12);
    }
  }

  if (isColumn(view.type) || isBar(view.type)) {
    const k = view.series.length;
    // Excel defaults: clustered gap 219% / overlap -27%; stacked gap 150% / overlap 100%.
    const barW = g === "standard" ? step / (k + (k - 1) * 0.27 + 2.19) : step / 2.5;
    const groupW = g === "standard" ? k * barW + (k - 1) * 0.27 * barW : barW;
    const thin = barW < 1;
    view.series.forEach((s, si) => {
      ctx.fillStyle = s.color;
      const offset = g === "standard" ? si * barW * 1.27 : 0;
      let lastPx = -1;
      segs[si].forEach((seg, i) => {
        if (!seg) return;
        const start = catStart(i) + (step - groupW) / 2 + offset;
        // More points than pixels: one bar per pixel column is enough.
        if (thin && Math.floor(start) === lastPx) return;
        lastPx = Math.floor(start);
        const a = valueAt(seg[0]);
        const b = valueAt(seg[1]);
        const w = Math.max(barW, 1);
        if (horizontal) ctx.fillRect(Math.min(a, b), start, Math.abs(b - a), w);
        else ctx.fillRect(start, Math.min(a, b), w, Math.abs(b - a));
      });
    });
    if (view.dataLabels && n <= 200) {
      ctx.fillStyle = TEXT;
      view.series.forEach((s, si) => {
        const offset = g === "standard" ? si * barW * 1.27 : 0;
        segs[si].forEach((seg, i) => {
          if (!seg) return;
          const mid = catStart(i) + (step - groupW) / 2 + offset + barW / 2;
          const end = valueAt(seg[1]);
          const inside = g !== "standard";
          const pos = inside ? (valueAt(seg[0]) + end) / 2 : end + (horizontal ? 4 : -8);
          ctx.textAlign = horizontal && !inside ? "left" : "center";
          if (horizontal) ctx.fillText(s.labels[i], pos, mid);
          else ctx.fillText(s.labels[i], mid, pos);
        });
      });
    }
    return;
  }

  // Lines and areas.
  const area = isArea(view.type);
  const pointX = (i: number) => (area && n > 1 ? plot.x + (i / (n - 1)) * plot.w : catMid(i));
  view.series.forEach((s, si) => drawLineSeries(ctx, s, segs[si], pointX, valueAt, area, view.type === "line-markers" && n <= plot.w / 3));
  if (view.dataLabels && n <= 200) {
    ctx.fillStyle = TEXT;
    ctx.textAlign = "center";
    view.series.forEach((s, si) =>
      segs[si].forEach((seg, i) => {
        if (seg) ctx.fillText(s.labels[i], pointX(i), valueAt(seg[1]) - 10);
      }),
    );
  }
}

function drawLineSeries(
  ctx: CanvasRenderingContext2D,
  s: SeriesView,
  segs: ([number, number] | null)[],
  pointX: (i: number) => number,
  valueAt: (v: number) => number,
  area: boolean,
  markers: boolean,
) {
  ctx.strokeStyle = s.color;
  ctx.fillStyle = s.color;
  ctx.lineWidth = 2.25;
  ctx.lineJoin = "round";
  ctx.lineCap = "round";
  if (area) {
    // Filled band between the stack base and the series; gaps count as zero.
    ctx.beginPath();
    segs.forEach((seg, i) => {
      const y = valueAt(seg ? seg[1] : 0);
      if (i === 0) ctx.moveTo(pointX(i), y);
      else ctx.lineTo(pointX(i), y);
    });
    for (let i = segs.length - 1; i >= 0; i--) ctx.lineTo(pointX(i), valueAt(segs[i] ? segs[i]![0] : 0));
    ctx.closePath();
    ctx.fill();
    return;
  }
  // Lines break at empty cells, like Excel's default "show empty cells as gaps".
  ctx.beginPath();
  let pen = false;
  let lastPx = -Infinity;
  let lo = Infinity;
  let hi = -Infinity;
  segs.forEach((seg, i) => {
    if (!seg) {
      pen = false;
      return;
    }
    const x = pointX(i);
    const y = valueAt(seg[1]);
    // More points than pixels: draw each pixel column's min and max.
    if (pen && x - lastPx < 1) {
      lo = Math.min(lo, y);
      hi = Math.max(hi, y);
      return;
    }
    if (pen && lo !== Infinity) {
      ctx.lineTo(lastPx, lo);
      ctx.lineTo(lastPx, hi);
    }
    lo = Infinity;
    hi = -Infinity;
    if (pen) ctx.lineTo(x, y);
    else ctx.moveTo(x, y);
    pen = true;
    lastPx = x;
  });
  ctx.stroke();
  if (markers) {
    segs.forEach((seg, i) => {
      if (!seg) return;
      ctx.beginPath();
      ctx.arc(pointX(i), valueAt(seg[1]), 3.5, 0, Math.PI * 2);
      ctx.fill();
    });
  }
}

function drawScatter(
  ctx: CanvasRenderingContext2D,
  view: ChartView,
  plot: Box,
  yAxis: AxisView,
  xAxis: AxisView,
  valueAt: (v: number) => number,
) {
  const xAt = (x: number) => plot.x + ((x - xAxis.min) / (xAxis.max - xAxis.min || 1)) * plot.w;
  // Vertical gridlines and x labels.
  ctx.strokeStyle = GRID;
  ctx.fillStyle = TEXT;
  ctx.textAlign = "center";
  xAxis.labels.forEach((label, i) => {
    const p = Math.round(xAt(xAxis.min + i * xAxis.step)) + 0.5;
    ctx.beginPath();
    ctx.moveTo(p, plot.y);
    ctx.lineTo(p, plot.y + plot.h);
    ctx.stroke();
    ctx.fillText(label, p, plot.y + plot.h + 12);
  });
  void yAxis;
  const lines = view.type === "scatter-lines";
  for (const s of view.series) {
    ctx.strokeStyle = s.color;
    ctx.fillStyle = s.color;
    ctx.lineWidth = 2.25;
    if (lines) {
      ctx.beginPath();
      let pen = false;
      s.values.forEach((v, i) => {
        const x = s.xs[i];
        if (v === null || x === null || x === undefined) {
          pen = false;
          return;
        }
        if (pen) ctx.lineTo(xAt(x), valueAt(v));
        else ctx.moveTo(xAt(x), valueAt(v));
        pen = true;
      });
      ctx.stroke();
    }
    const many = s.values.length > 5000;
    s.values.forEach((v, i) => {
      const x = s.xs[i];
      if (v === null || x === null || x === undefined) return;
      if (many) ctx.fillRect(xAt(x) - 1, valueAt(v) - 1, 2, 2);
      else {
        ctx.beginPath();
        ctx.arc(xAt(x), valueAt(v), 3.5, 0, Math.PI * 2);
        ctx.fill();
      }
    });
    if (view.dataLabels && s.values.length <= 200) {
      ctx.fillStyle = TEXT;
      s.values.forEach((v, i) => {
        const x = s.xs[i];
        if (v !== null && x !== null && x !== undefined) ctx.fillText(s.labels[i], xAt(x), valueAt(v) - 10);
      });
    }
  }
}
