// Draws a shape (Insert → Shapes) on a canvas: the DrawingML preset geometry
// with Excel's default proportions, its fill and outline, an arrowhead for
// arrows, and its text wrapped inside Excel's text insets. Pure; used by the
// drawing layer on screen and for printing.

import type { ShapeInfo } from "./types";

const FONT = "Calibri, Carlito, 'Segoe UI', Arial, sans-serif";
/** Excel's text box insets: 0.1" left/right, 0.05" top/bottom. */
const INSET_X = 9.6;
const INSET_Y = 4.8;

type Path = (p: Path2D, w: number, h: number) => void;

const poly = (pts: [number, number][]): Path => (p, w, h) => {
  pts.forEach(([x, y], i) => (i ? p.lineTo(x * w, y * h) : p.moveTo(x * w, y * h)));
  p.closePath();
};

function star(points: number, inner: number): Path {
  return (p, w, h) => {
    for (let i = 0; i < points * 2; i++) {
      const a = -Math.PI / 2 + (i * Math.PI) / points;
      const r = i % 2 ? inner : 1;
      const x = w / 2 + (Math.cos(a) * r * w) / 2;
      const y = h / 2 + (Math.sin(a) * r * h) / 2;
      if (i) p.lineTo(x, y);
      else p.moveTo(x, y);
    }
    p.closePath();
  };
}

function roundRect(p: Path2D, x: number, y: number, w: number, h: number, r: number) {
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  p.moveTo(x + r, y);
  p.arcTo(x + w, y, x + w, y + h, r);
  p.arcTo(x + w, y + h, x, y + h, r);
  p.arcTo(x, y + h, x, y, r);
  p.arcTo(x, y, x + w, y, r);
  p.closePath();
}

/** Block arrow pointing right, in a w x h box (rotate for the others). */
function arrowRight(p: Path2D, w: number, h: number) {
  const head = Math.min(w, h) * 0.5;
  const shaft = h * 0.25;
  p.moveTo(0, shaft);
  p.lineTo(w - head, shaft);
  p.lineTo(w - head, 0);
  p.lineTo(w, h / 2);
  p.lineTo(w - head, h);
  p.lineTo(w - head, h - shaft);
  p.lineTo(0, h - shaft);
  p.closePath();
}

/** A wedge from the box to a point below-left, for callouts. */
function wedge(p: Path2D, w: number, h: number, inset: number) {
  const tipX = w * 0.29;
  const tipY = h * 1.125;
  p.moveTo(w * 0.22, h - inset);
  p.lineTo(tipX, tipY);
  p.lineTo(w * 0.42, h - inset);
}

const PATHS: Record<string, Path> = {
  rect: (p, w, h) => p.rect(0, 0, w, h),
  flowChartProcess: (p, w, h) => p.rect(0, 0, w, h),
  roundRect: (p, w, h) => roundRect(p, 0, 0, w, h, Math.min(w, h) * 0.1667),
  snip1Rect: (p, w, h) => {
    const s = Math.min(w, h) * 0.1667;
    poly([[0, 0], [(w - s) / w, 0], [1, s / h], [1, 1], [0, 1]])(p, w, h);
  },
  ellipse: (p, w, h) => p.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2),
  triangle: poly([[0.5, 0], [1, 1], [0, 1]]),
  rtTriangle: poly([[0, 0], [1, 1], [0, 1]]),
  parallelogram: (p, w, h) => {
    const o = Math.min(w, h) * 0.25;
    poly([[o / w, 0], [1, 0], [1 - o / w, 1], [0, 1]])(p, w, h);
  },
  flowChartInputOutput: poly([[0.2, 0], [1, 0], [0.8, 1], [0, 1]]),
  trapezoid: (p, w, h) => {
    const o = Math.min(w, h) * 0.25;
    poly([[o / w, 0], [1 - o / w, 0], [1, 1], [0, 1]])(p, w, h);
  },
  diamond: poly([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]),
  flowChartDecision: poly([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]),
  pentagon: poly([[0.5, 0], [1, 0.382], [0.809, 1], [0.191, 1], [0, 0.382]]),
  hexagon: (p, w, h) => {
    const o = Math.min(w, h) * 0.25;
    poly([[o / w, 0], [1 - o / w, 0], [1, 0.5], [1 - o / w, 1], [o / w, 1], [0, 0.5]])(p, w, h);
  },
  octagon: (p, w, h) => {
    const o = Math.min(w, h) * 0.2929;
    const x = o / w;
    const y = o / h;
    poly([[x, 0], [1 - x, 0], [1, y], [1, 1 - y], [1 - x, 1], [x, 1], [0, 1 - y], [0, y]])(p, w, h);
  },
  plus: (p, w, h) => {
    const o = Math.min(w, h) * 0.25;
    const x = o / w;
    const y = o / h;
    poly([[x, 0], [1 - x, 0], [1 - x, y], [1, y], [1, 1 - y], [1 - x, 1 - y], [1 - x, 1], [x, 1], [x, 1 - y], [0, 1 - y], [0, y], [x, y]])(p, w, h);
  },
  donut: (p, w, h) => {
    const t = Math.min(w, h) * 0.25;
    p.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
    p.moveTo(w - t, h / 2);
    p.ellipse(w / 2, h / 2, w / 2 - t, h / 2 - t, 0, 0, Math.PI * 2, true);
  },
  heart: (p, w, h) => {
    p.moveTo(w / 2, h * 0.25);
    p.bezierCurveTo(w * 0.5, -h * 0.05, w * -0.05, h * 0.05, w * 0.05, h * 0.4);
    p.bezierCurveTo(w * 0.12, h * 0.65, w * 0.4, h * 0.8, w / 2, h);
    p.bezierCurveTo(w * 0.6, h * 0.8, w * 0.88, h * 0.65, w * 0.95, h * 0.4);
    p.bezierCurveTo(w * 1.05, h * 0.05, w * 0.5, -h * 0.05, w / 2, h * 0.25);
    p.closePath();
  },
  smileyFace: (p, w, h) => p.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2),
  can: (p, w, h) => {
    const r = Math.min(h * 0.25, w * 0.25) / 2;
    p.moveTo(0, r);
    p.ellipse(w / 2, r, w / 2, r, 0, Math.PI, 0);
    p.lineTo(w, h - r);
    p.ellipse(w / 2, h - r, w / 2, r, 0, 0, Math.PI);
    p.closePath();
  },
  cube: (p, w, h) => {
    const d = Math.min(w, h) * 0.25;
    poly([[0, d / h], [d / w, 0], [1, 0], [1, 1 - d / h], [1 - d / w, 1], [0, 1]])(p, w, h);
  },
  rightArrow: (p, w, h) => arrowRight(p, w, h),
  leftArrow: (p, w, h) => {
    const q = new Path2D();
    arrowRight(q, w, h);
    p.addPath(q, new DOMMatrix([-1, 0, 0, 1, w, 0]));
  },
  upArrow: (p, w, h) => {
    const q = new Path2D();
    arrowRight(q, h, w);
    p.addPath(q, new DOMMatrix([0, -1, 1, 0, 0, h]));
  },
  downArrow: (p, w, h) => {
    const q = new Path2D();
    arrowRight(q, h, w);
    p.addPath(q, new DOMMatrix([0, 1, 1, 0, 0, 0]));
  },
  leftRightArrow: (p, w, h) => {
    const head = Math.min(w / 2, h) * 0.5;
    const s = h * 0.25;
    p.moveTo(0, h / 2);
    p.lineTo(head, 0);
    p.lineTo(head, s);
    p.lineTo(w - head, s);
    p.lineTo(w - head, 0);
    p.lineTo(w, h / 2);
    p.lineTo(w - head, h);
    p.lineTo(w - head, h - s);
    p.lineTo(head, h - s);
    p.lineTo(head, h);
    p.closePath();
  },
  chevron: (p, w, h) => {
    const d = Math.min(w, h) * 0.5;
    poly([[0, 0], [1 - d / w, 0], [1, 0.5], [1 - d / w, 1], [0, 1], [d / w, 0.5]])(p, w, h);
  },
  homePlate: (p, w, h) => {
    const d = Math.min(w, h) * 0.5;
    poly([[0, 0], [1 - d / w, 0], [1, 0.5], [1 - d / w, 1], [0, 1]])(p, w, h);
  },
  star5: star(5, 0.382),
  star6: star(6, 0.577),
  star8: star(8, 0.7),
  irregularSeal1: poly([
    [0.39, 0.25], [0.5, 0], [0.56, 0.22], [0.79, 0.07], [0.71, 0.3], [0.97, 0.3], [0.79, 0.45], [1, 0.62], [0.75, 0.64],
    [0.83, 0.9], [0.6, 0.76], [0.55, 1], [0.44, 0.78], [0.23, 0.95], [0.27, 0.71], [0, 0.73], [0.17, 0.53], [0.01, 0.38], [0.24, 0.36], [0.13, 0.12],
  ]),
  wedgeRectCallout: (p, w, h) => {
    p.moveTo(0, 0);
    p.lineTo(w, 0);
    p.lineTo(w, h);
    p.lineTo(w * 0.42, h);
    p.lineTo(w * 0.29, h * 1.125);
    p.lineTo(w * 0.22, h);
    p.lineTo(0, h);
    p.closePath();
  },
  wedgeRoundRectCallout: (p, w, h) => {
    roundRect(p, 0, 0, w, h, Math.min(w, h) * 0.1667);
    wedge(p, w, h, 1);
    p.closePath();
  },
  wedgeEllipseCallout: (p, w, h) => {
    p.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
    p.moveTo(w * 0.3, h * 0.92);
    p.lineTo(w * 0.29, h * 1.125);
    p.lineTo(w * 0.45, h * 0.99);
    p.closePath();
  },
  cloudCallout: (p, w, h) => {
    const bumps: [number, number, number][] = [
      [0.25, 0.35, 0.22], [0.45, 0.22, 0.24], [0.68, 0.25, 0.22], [0.82, 0.48, 0.2], [0.65, 0.72, 0.22], [0.4, 0.75, 0.22], [0.18, 0.58, 0.2],
    ];
    for (const [x, y, r] of bumps) {
      p.moveTo(x * w + r * w, y * h);
      p.ellipse(x * w, y * h, r * w, r * h * 1.1, 0, 0, Math.PI * 2);
    }
    for (const [x, y, r] of [[0.2, 0.95, 0.04], [0.12, 1.05, 0.025]] as const) {
      p.moveTo(x * w + r * w, y * h);
      p.ellipse(x * w, y * h, r * w, r * h, 0, 0, Math.PI * 2);
    }
  },
  flowChartTerminator: (p, w, h) => roundRect(p, 0, 0, w, h, h / 2),
  flowChartDocument: (p, w, h) => {
    p.moveTo(0, 0);
    p.lineTo(w, 0);
    p.lineTo(w, h * 0.83);
    p.bezierCurveTo(w * 0.75, h * 0.7, w * 0.5, h * 1.1, 0, h * 0.95);
    p.closePath();
  },
};

/** Draw `shape` filling a w x h box at the origin. */
export function drawShape(ctx: CanvasRenderingContext2D, shape: ShapeInfo, w: number, h: number) {
  ctx.save();
  const lineW = Math.max(0.5, shape.lineWidth * (4 / 3));
  if (shape.preset === "line" || shape.preset === "straightConnector1") {
    const [x1, y1, x2, y2] = shape.flipV ? [0, h, w, 0] : [0, 0, w, h];
    ctx.strokeStyle = shape.line ?? "transparent";
    ctx.lineWidth = lineW;
    ctx.lineCap = "round";
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    if (shape.arrowEnd && shape.line) {
      const a = Math.atan2(y2 - y1, x2 - x1);
      const size = Math.max(6, lineW * 3);
      ctx.fillStyle = shape.line;
      ctx.beginPath();
      ctx.moveTo(x2, y2);
      ctx.lineTo(x2 - size * Math.cos(a - 0.45), y2 - size * Math.sin(a - 0.45));
      ctx.lineTo(x2 - size * Math.cos(a + 0.45), y2 - size * Math.sin(a + 0.45));
      ctx.closePath();
      ctx.fill();
    }
    ctx.restore();
    return;
  }
  // Keep the outline inside the box so it isn't clipped.
  const inset = shape.line ? lineW / 2 : 0;
  const bw = Math.max(1, w - 2 * inset);
  const bh = Math.max(1, h - 2 * inset);
  const path = new Path2D();
  const build = PATHS[shape.preset] ?? PATHS.rect;
  // Callout pointers reach below the box; shrink the body so they fit.
  const callout = shape.preset.startsWith("wedge") || shape.preset === "cloudCallout";
  build(path, bw, callout ? bh / 1.125 : bh);
  ctx.translate(inset, inset);
  if (shape.fill) {
    ctx.fillStyle = shape.fill;
    ctx.fill(path, "evenodd");
  }
  if (shape.line) {
    ctx.strokeStyle = shape.line;
    ctx.lineWidth = lineW;
    ctx.lineJoin = "round";
    ctx.stroke(path);
  }
  if (shape.preset === "smileyFace" && shape.line) {
    // Eyes and smile.
    ctx.fillStyle = shape.line;
    for (const ex of [0.35, 0.65]) {
      ctx.beginPath();
      ctx.ellipse(bw * ex, bh * 0.38, bw * 0.05, bh * 0.05, 0, 0, Math.PI * 2);
      ctx.fill();
    }
    ctx.beginPath();
    ctx.ellipse(bw / 2, bh * 0.55, bw * 0.25, bh * 0.18, 0, 0.15 * Math.PI, 0.85 * Math.PI);
    ctx.stroke();
  }
  if (shape.preset === "cube" && shape.line) {
    const d = Math.min(bw, bh) * 0.25;
    ctx.beginPath();
    ctx.moveTo(0, d);
    ctx.lineTo(bw - d, d);
    ctx.lineTo(bw, 0);
    ctx.moveTo(bw - d, d);
    ctx.lineTo(bw - d, bh);
    ctx.stroke();
  }
  if (shape.preset === "can" && shape.line) {
    const r = Math.min(bh * 0.25, bw * 0.25) / 2;
    ctx.beginPath();
    ctx.ellipse(bw / 2, r, bw / 2, r, 0, 0, Math.PI);
    ctx.stroke();
  }
  ctx.translate(-inset, -inset);
  if (shape.text) drawText(ctx, shape, w, callout ? h / 1.125 : h);
  ctx.restore();
}

/** Lines of `text` that fit `width`, breaking at spaces (and inside long words). */
export function wrapText(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split("\n")) {
    let line = "";
    for (const word of para.split(" ")) {
      const next = line ? `${line} ${word}` : word;
      if (ctx.measureText(next).width <= width || !line) {
        line = next;
      } else {
        out.push(line);
        line = word;
      }
      // A single word wider than the box breaks by characters.
      while (ctx.measureText(line).width > width && line.length > 1) {
        let cut = line.length - 1;
        while (cut > 1 && ctx.measureText(line.slice(0, cut)).width > width) cut--;
        out.push(line.slice(0, cut));
        line = line.slice(cut);
      }
    }
    out.push(line);
  }
  return out;
}

function drawText(ctx: CanvasRenderingContext2D, s: ShapeInfo, w: number, h: number) {
  const px = s.fontSize * (4 / 3);
  ctx.font = `${s.italic ? "italic " : ""}${s.bold ? "bold " : ""}${px}px ${FONT}`;
  ctx.fillStyle = s.textColor;
  ctx.textBaseline = "top";
  const width = Math.max(1, w - 2 * INSET_X);
  const lines = wrapText(ctx, s.text, width);
  const lineH = px * 1.2;
  const total = lines.length * lineH;
  let y = s.textTop ? INSET_Y : (h - total) / 2;
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, 0, w, h);
  ctx.clip();
  for (const line of lines) {
    const tw = ctx.measureText(line).width;
    const x = s.align === "left" ? INSET_X : s.align === "right" ? w - INSET_X - tw : (w - tw) / 2;
    ctx.fillText(line, x, y);
    y += lineH;
  }
  ctx.restore();
}
