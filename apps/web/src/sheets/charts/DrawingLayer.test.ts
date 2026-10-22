import { describe, expect, it } from "vitest";
import { ColumnLayout, RowLayout } from "../layout";
import { MergeIndex } from "../merges";
import type { SheetGeometry } from "../paint";
import { anchorRect, rectAnchor } from "./DrawingLayer";

const geo = (cols = new ColumnLayout(), rows = new RowLayout()): SheetGeometry => ({
  cols,
  rows,
  merges: new MergeIndex([]),
  freeze: { rows: 0, cols: 0, top: 0, left: 0 },
});

describe("drawing anchors", () => {
  it("round-trips pixels through cell anchors", () => {
    const g = geo();
    const a = rectAnchor(g, { x: 200, y: 50, w: 480, h: 288 });
    // 200px = column D (192) + 8px; 50px = row 3 (40) + 10px.
    expect(a.slice(0, 4)).toEqual([2, 3, 10, 8]);
    expect(anchorRect(g, a)).toEqual({ x: 200, y: 50, w: 480, h: 288 });
  });

  it("follows column resizes", () => {
    const a = rectAnchor(geo(), { x: 200, y: 0, w: 100, h: 100 });
    const wider = geo(new ColumnLayout([0, 164]));
    expect(anchorRect(wider, a).x).toBe(300);
  });
});
