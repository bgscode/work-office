// Hand-drawn ribbon icons in the spirit of Microsoft Office's monoline set:
// dark grey strokes with a few colour accents. All are drawn on a 16×16 grid
// and scale cleanly to the 32px "large" ribbon size.
//
// Use by name from a RibbonCommand (`icon: "bold"`) or as a component
// (`<Icon name="bold" />`). Unknown names render nothing.

import type { ReactElement } from "react";

const INK = "#444444";
const GREEN = "#217346";
const BLUE = "#2b78d0";
const RED = "#c43e1c";
const GOLD = "#c19c00";

const s = { fill: "none", stroke: INK, strokeWidth: 1, strokeLinecap: "round", strokeLinejoin: "round" } as const;

type Paths = () => ReactElement;

const ICONS: Record<string, Paths> = {
  paste: () => (
    <>
      <rect x="2.5" y="2.5" width="9" height="12" rx="1" {...s} fill="#fff" />
      <rect x="5" y="1.5" width="4" height="2.5" rx="0.6" {...s} fill={GOLD} stroke={GOLD} />
      <rect x="7.5" y="7.5" width="6" height="7" {...s} fill="#fff" />
      <path d="M9 10h3M9 12h3" {...s} />
    </>
  ),
  cut: () => (
    <>
      <circle cx="4.5" cy="12" r="2" {...s} stroke={BLUE} />
      <circle cx="11.5" cy="12" r="2" {...s} stroke={BLUE} />
      <path d="M5.8 10.5 11 2M10.2 10.5 5 2" {...s} />
    </>
  ),
  copy: () => (
    <>
      <rect x="2.5" y="2.5" width="7" height="9" rx="0.5" {...s} />
      <rect x="6.5" y="5.5" width="7" height="9" rx="0.5" {...s} fill="#fff" />
    </>
  ),
  formatPainter: () => (
    <>
      <rect x="3.5" y="1.5" width="10" height="4" rx="0.5" {...s} fill={GOLD} stroke={INK} />
      <path d="M8.5 5.5v2.5h-3v2" {...s} />
      <rect x="4.5" y="10" width="2" height="5" rx="0.5" {...s} />
    </>
  ),
  pasteSpecial: () => (
    <>
      <rect x="2.5" y="2.5" width="9" height="12" rx="1" {...s} fill="#fff" />
      <rect x="5" y="1.5" width="4" height="2.5" rx="0.6" {...s} fill={GOLD} stroke={GOLD} />
      <path d="M5 8h4M5 10h4M5 12h2" {...s} stroke={BLUE} />
    </>
  ),
  undo: () => <path d="M5 4 2 7l3 3M2.5 7H10a3.5 3.5 0 0 1 0 7H7" {...s} stroke={BLUE} strokeWidth={1.3} />,
  redo: () => <path d="m11 4 3 3-3 3M13.5 7H6a3.5 3.5 0 0 0 0 7h3" {...s} stroke={BLUE} strokeWidth={1.3} />,
  bold: () => <path d="M4.5 2.5h4.25a2.75 2.75 0 0 1 0 5.5H4.5zm0 5.5h5a3 3 0 0 1 0 6h-5z" {...s} strokeWidth={1.6} />,
  italic: () => <path d="M7 2.5h5M4 13.5h5M9.5 2.5l-3 11" {...s} strokeWidth={1.3} />,
  underline: () => (
    <>
      <path d="M4.5 2.5v5a3.5 3.5 0 0 0 7 0v-5" {...s} strokeWidth={1.3} />
      <path d="M3.5 14.5h9" {...s} />
    </>
  ),
  strike: () => (
    <>
      <path d="M11 4.5c-.5-1.3-1.7-2-3.2-2-1.8 0-3.1 1-3.1 2.4 0 1.5 1.3 2 3.3 2.5M5 11.5c.5 1.3 1.8 2 3.3 2 1.9 0 3.2-1 3.2-2.5 0-.8-.3-1.4-.9-1.8" {...s} />
      <path d="M2 8.5h12" {...s} stroke={RED} />
    </>
  ),
  fontColor: () => <path d="M4 11 8 2l4 9M5.4 8h5.2" {...s} strokeWidth={1.3} />,
  fillColor: () => (
    <>
      <path d="m7 1.5 5.5 5.5-4.5 4.5L2.5 6z" {...s} />
      <path d="M13.5 9.5s1 1.4 1 2a1 1 0 0 1-2 0c0-.6 1-2 1-2z" {...s} fill={INK} />
    </>
  ),
  borders: () => (
    <>
      <rect x="2.5" y="2.5" width="11" height="11" {...s} strokeDasharray="1 1.2" />
      <path d="M2.5 13.5h11" {...s} strokeWidth={1.4} />
    </>
  ),
  alignLeft: () => <path d="M2 3.5h12M2 6.5h8M2 9.5h12M2 12.5h8" {...s} />,
  alignCenter: () => <path d="M2 3.5h12M4 6.5h8M2 9.5h12M4 12.5h8" {...s} />,
  alignRight: () => <path d="M2 3.5h12M6 6.5h8M2 9.5h12M6 12.5h8" {...s} />,
  alignTop: () => (
    <>
      <path d="M2 2.5h12" {...s} />
      <path d="M8 13.5v-9M5.5 7 8 4.5 10.5 7" {...s} stroke={BLUE} />
    </>
  ),
  alignMiddle: () => (
    <>
      <path d="M2 8h12" {...s} />
      <path d="M8 1.5v4M6 3.5 8 5.5l2-2M8 14.5v-4M6 12.5l2-2 2 2" {...s} stroke={BLUE} />
    </>
  ),
  alignBottom: () => (
    <>
      <path d="M2 13.5h12" {...s} />
      <path d="M8 2.5v9M5.5 9 8 11.5 10.5 9" {...s} stroke={BLUE} />
    </>
  ),
  wrap: () => (
    <>
      <path d="M2 3.5h12M2 7.5h10a2 2 0 0 1 0 4H8.5M2 11.5h4" {...s} />
      <path d="m10 10-1.5 1.5L10 13" {...s} stroke={BLUE} />
    </>
  ),
  merge: () => (
    <>
      <rect x="1.5" y="3.5" width="13" height="9" {...s} />
      <path d="M4 8h8M5.5 6.5 4 8l1.5 1.5M10.5 6.5 12 8l-1.5 1.5" {...s} stroke={BLUE} />
    </>
  ),
  indentDecrease: () => <path d="M7 3.5h7M7 6.5h7M7 9.5h7M2 12.5h12M4.5 5 2.5 7l2 2" {...s} />,
  indentIncrease: () => <path d="M7 3.5h7M7 6.5h7M7 9.5h7M2 12.5h12M2.5 5l2 2-2 2" {...s} />,
  currency: () => (
    <>
      <rect x="1.5" y="4.5" width="13" height="8" rx="1" {...s} stroke={GREEN} />
      <circle cx="8" cy="8.5" r="2" {...s} stroke={GREEN} />
    </>
  ),
  percent: () => (
    <>
      <circle cx="4.5" cy="4.5" r="1.8" {...s} />
      <circle cx="11.5" cy="11.5" r="1.8" {...s} />
      <path d="M12.5 3 3.5 13" {...s} />
    </>
  ),
  comma: () => <text x="3" y="13" fontSize="12" fontWeight="700" fill={INK} fontFamily="Segoe UI, Arial">,</text>,
  decimalIncrease: () => (
    <>
      <text x="0.5" y="8" fontSize="6.5" fill={INK} fontFamily="Segoe UI, Arial">.0</text>
      <text x="6.5" y="8" fontSize="6.5" fill={INK} fontFamily="Segoe UI, Arial">.00</text>
      <path d="M3 11.5h7M8 9.5l2 2-2 2" {...s} stroke={BLUE} />
    </>
  ),
  decimalDecrease: () => (
    <>
      <text x="0.5" y="8" fontSize="6.5" fill={INK} fontFamily="Segoe UI, Arial">.00</text>
      <text x="10" y="8" fontSize="6.5" fill={INK} fontFamily="Segoe UI, Arial">.0</text>
      <path d="M13 11.5H6M8 9.5l-2 2 2 2" {...s} stroke={BLUE} />
    </>
  ),
  numberFormat: () => (
    <>
      <rect x="1.5" y="3.5" width="13" height="9" rx="1" {...s} />
      <text x="3" y="10.5" fontSize="6" fill={INK} fontFamily="Segoe UI, Arial">123</text>
    </>
  ),
  insertCells: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" {...s} />
      <path d="M1.5 6.5h13M1.5 10h13M6 2.5v11" {...s} />
      <path d="M10 6.5v3.5" {...s} />
      <rect x="6" y="6.5" width="8.5" height="3.5" fill={GREEN} opacity="0.35" />
    </>
  ),
  deleteCells: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" {...s} />
      <path d="M1.5 6.5h13M1.5 10h13M6 2.5v11" {...s} />
      <path d="m8 7 4 2.5M12 7l-4 2.5" {...s} stroke={RED} />
    </>
  ),
  formatCells: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" {...s} />
      <path d="M1.5 6.5h13M1.5 10h13M6 2.5v11" {...s} />
      <rect x="6" y="6.5" width="8.5" height="3.5" fill={BLUE} opacity="0.35" />
    </>
  ),
  insertRowAbove: () => (
    <>
      <rect x="1.5" y="8.5" width="13" height="5" {...s} />
      <path d="M8 1.5v5M5.5 4h5" {...s} stroke={GREEN} strokeWidth={1.4} />
    </>
  ),
  insertRowBelow: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="5" {...s} />
      <path d="M8 9.5v5M5.5 12h5" {...s} stroke={GREEN} strokeWidth={1.4} />
    </>
  ),
  insertColumnLeft: () => (
    <>
      <rect x="8.5" y="1.5" width="5" height="13" {...s} />
      <path d="M4 5.5v5M1.5 8h5" {...s} stroke={GREEN} strokeWidth={1.4} />
    </>
  ),
  insertColumnRight: () => (
    <>
      <rect x="2.5" y="1.5" width="5" height="13" {...s} />
      <path d="M12 5.5v5M9.5 8h5" {...s} stroke={GREEN} strokeWidth={1.4} />
    </>
  ),
  deleteRows: () => (
    <>
      <rect x="1.5" y="5.5" width="13" height="5" {...s} />
      <path d="m6 6.5 4 3M10 6.5l-4 3" {...s} stroke={RED} />
    </>
  ),
  deleteColumns: () => (
    <>
      <rect x="5.5" y="1.5" width="5" height="13" {...s} />
      <path d="m6.5 6 3 4M9.5 6l-3 4" {...s} stroke={RED} />
    </>
  ),
  insertSheet: () => (
    <>
      <path d="M3.5 1.5h6l3 3v10h-9z" {...s} fill="#fff" />
      <path d="M9.5 1.5v3h3" {...s} />
      <path d="M8 7.5v5M5.5 10h5" {...s} stroke={GREEN} strokeWidth={1.4} />
    </>
  ),
  clear: () => (
    <>
      <path d="m9 2.5 4.5 4.5-6 6H4l-1.5-1.5z" {...s} />
      <path d="m5.5 6.5 4.5 4.5M7.5 13.5h7" {...s} stroke={RED} />
    </>
  ),
  sortAsc: () => (
    <>
      <text x="1" y="7" fontSize="6.5" fontWeight="700" fill={BLUE} fontFamily="Segoe UI, Arial">A</text>
      <text x="1" y="15" fontSize="6.5" fontWeight="700" fill={INK} fontFamily="Segoe UI, Arial">Z</text>
      <path d="M11.5 2v12M9 11.5l2.5 2.5 2.5-2.5" {...s} />
    </>
  ),
  sortDesc: () => (
    <>
      <text x="1" y="7" fontSize="6.5" fontWeight="700" fill={INK} fontFamily="Segoe UI, Arial">Z</text>
      <text x="1" y="15" fontSize="6.5" fontWeight="700" fill={BLUE} fontFamily="Segoe UI, Arial">A</text>
      <path d="M11.5 2v12M9 11.5l2.5 2.5 2.5-2.5" {...s} />
    </>
  ),
  sortFilter: () => (
    <>
      <text x="0.5" y="7" fontSize="6.5" fontWeight="700" fill={BLUE} fontFamily="Segoe UI, Arial">A</text>
      <text x="0.5" y="15" fontSize="6.5" fontWeight="700" fill={INK} fontFamily="Segoe UI, Arial">Z</text>
      <path d="M7.5 3.5h7l-2.75 4v4l-1.5 1v-5z" {...s} />
    </>
  ),
  filter: () => <path d="M2.5 2.5h11l-4.25 5.5v5l-2.5 1.5V8z" {...s} />,
  find: () => (
    <>
      <circle cx="6.5" cy="6.5" r="4" {...s} />
      <path d="m9.5 9.5 4.5 4.5" {...s} strokeWidth={1.6} />
    </>
  ),
  replace: () => (
    <>
      <path d="M2.5 5.5a4 4 0 0 1 7-2.5l1 1M2.5 2.5v3h3" {...s} />
      <path d="M13.5 10.5a4 4 0 0 1-7 2.5l-1-1M13.5 13.5v-3h-3" {...s} stroke={BLUE} />
    </>
  ),
  goTo: () => <path d="M2 8h10M8.5 4.5 12 8l-3.5 3.5M14 3v10" {...s} />,
  autoSum: () => <path d="M12.5 3.5v-1h-9l5 5.5-5 5.5h9v-1" {...s} strokeWidth={1.3} />,
  fill: () => (
    <>
      <rect x="2.5" y="1.5" width="11" height="4" {...s} fill="#e8f2ec" />
      <path d="M8 6.5v7M5.5 11 8 13.5 10.5 11" {...s} stroke={BLUE} />
    </>
  ),
  zoom: () => (
    <>
      <circle cx="6.5" cy="6.5" r="4.5" {...s} />
      <path d="m10 10 4 4" {...s} strokeWidth={1.6} />
      <path d="M4.5 6.5h4M6.5 4.5v4" {...s} stroke={BLUE} />
    </>
  ),
  zoomIn: () => (
    <>
      <circle cx="6.5" cy="6.5" r="4.5" {...s} />
      <path d="m10 10 4 4M4.5 6.5h4M6.5 4.5v4" {...s} />
    </>
  ),
  zoomOut: () => (
    <>
      <circle cx="6.5" cy="6.5" r="4.5" {...s} />
      <path d="m10 10 4 4M4.5 6.5h4" {...s} />
    </>
  ),
  zoom100: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" {...s} />
      <text x="2.6" y="10.6" fontSize="5.6" fontWeight="700" fill={BLUE} fontFamily="Segoe UI, Arial">100</text>
    </>
  ),
  zoomSelection: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" {...s} />
      <rect x="4.5" y="5.5" width="7" height="5" {...s} stroke={GREEN} strokeWidth={1.4} />
    </>
  ),
  function: () => (
    <text x="1.5" y="12.5" fontSize="11" fontStyle="italic" fontFamily="Cambria, Georgia, serif" fill={INK}>
      fx
    </text>
  ),
  newDocument: () => (
    <>
      <path d="M3.5 1.5h6l3 3v10h-9z" {...s} fill="#fff" />
      <path d="M9.5 1.5v3h3" {...s} />
    </>
  ),
  open: () => <path d="M1.5 3.5h4l1.5 1.5h6.5v2M1.5 3.5v10h10.5l2.5-6.5H4L1.5 13.5" {...s} stroke={GOLD} />,
  save: () => (
    <>
      <path d="M2.5 2.5h9l2 2v9h-11z" {...s} />
      <rect x="5" y="2.5" width="5" height="3.5" {...s} />
      <rect x="4.5" y="9" width="7" height="4.5" {...s} />
    </>
  ),
  download: () => (
    <>
      <path d="M8 2v8M4.5 6.5 8 10l3.5-3.5" {...s} stroke={GREEN} strokeWidth={1.3} />
      <path d="M2.5 11v2.5h11V11" {...s} />
    </>
  ),
  excelFile: () => (
    <>
      <path d="M4.5 1.5h6l3 3v10h-9z" {...s} fill="#fff" />
      <rect x="1.5" y="5" width="7" height="7" rx="0.8" fill={GREEN} />
      <path d="m3.3 6.8 3.4 3.4M6.7 6.8l-3.4 3.4" stroke="#fff" strokeWidth={1.2} />
    </>
  ),
  csvFile: () => (
    <>
      <path d="M4.5 1.5h6l3 3v10h-9z" {...s} fill="#fff" />
      <path d="M6.5 7h5M6.5 9h5M6.5 11h5" {...s} />
    </>
  ),
  table: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" {...s} />
      <path d="M1.5 6h13M1.5 9.5h13M6 2.5v11M10 2.5v11" {...s} />
      <rect x="1.5" y="2.5" width="13" height="3.5" fill={GREEN} opacity="0.4" />
    </>
  ),
  chartColumn: () => (
    <>
      <path d="M1.5 14.5h13" {...s} />
      <rect x="2.5" y="7" width="3" height="7.5" fill={BLUE} />
      <rect x="6.5" y="3" width="3" height="11.5" fill={GREEN} />
      <rect x="10.5" y="9" width="3" height="5.5" fill={GOLD} />
    </>
  ),
  chartBar: () => (
    <>
      <path d="M1.5 1.5v13" {...s} />
      <rect x="1.5" y="2.5" width="8" height="3" fill={BLUE} />
      <rect x="1.5" y="6.5" width="12" height="3" fill={GREEN} />
      <rect x="1.5" y="10.5" width="5" height="3" fill={GOLD} />
    </>
  ),
  chartLine: () => (
    <>
      <path d="M1.5 1.5v13h13" {...s} />
      <path d="m3 11 3.5-4 3 2.5L14 3.5" {...s} stroke={BLUE} strokeWidth={1.4} />
    </>
  ),
  chartPie: () => (
    <>
      <circle cx="8" cy="8" r="6" fill={BLUE} />
      <path d="M8 8V2a6 6 0 0 1 6 6z" fill={GOLD} />
      <path d="M8 8h6a6 6 0 0 1-2 4.5z" fill={GREEN} />
    </>
  ),
  chartScatter: () => (
    <>
      <path d="M1.5 1.5v13h13" {...s} />
      <circle cx="5" cy="10.5" r="1.3" fill={BLUE} />
      <circle cx="7.5" cy="7" r="1.3" fill={BLUE} />
      <circle cx="10.5" cy="8.5" r="1.3" fill={GOLD} />
      <circle cx="12.5" cy="4" r="1.3" fill={GOLD} />
    </>
  ),
  picture: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="0.5" {...s} />
      <circle cx="5.5" cy="6" r="1.5" fill={GOLD} />
      <path d="M2 12.5l4-4 3 3 2-2 3.5 3.5" {...s} stroke={GREEN} />
    </>
  ),
  chartElement: () => (
    <>
      <rect x="1.5" y="2.5" width="11" height="11" {...s} />
      <rect x="3.5" y="8" width="2" height="4" fill={BLUE} />
      <rect x="7" y="5.5" width="2" height="6.5" fill={BLUE} />
      <path d="M12.5 1v5M10 3.5h5" {...s} stroke={GREEN} strokeWidth={1.5} />
    </>
  ),
  chartColors: () => (
    <>
      <rect x="2" y="2" width="5.5" height="5.5" fill={BLUE} />
      <rect x="8.5" y="2" width="5.5" height="5.5" fill={GOLD} />
      <rect x="2" y="8.5" width="5.5" height="5.5" fill={GREEN} />
      <rect x="8.5" y="8.5" width="5.5" height="5.5" fill={RED} />
    </>
  ),
  switchRowCol: () => (
    <>
      <rect x="1.5" y="1.5" width="6" height="6" {...s} />
      <rect x="8.5" y="8.5" width="6" height="6" {...s} fill={BLUE} fillOpacity={0.3} />
      <path d="M10 2.5h2.5V5M12.5 2.5L9.5 5.5M6 13.5H3.5V11M3.5 13.5l3-3" {...s} stroke={GREEN} />
    </>
  ),
  bringFront: () => (
    <>
      <rect x="1.5" y="1.5" width="8" height="8" {...s} />
      <rect x="6.5" y="6.5" width="8" height="8" {...s} fill={BLUE} />
    </>
  ),
  sendBack: () => (
    <>
      <rect x="6.5" y="6.5" width="8" height="8" {...s} fill={BLUE} />
      <rect x="1.5" y="1.5" width="8" height="8" {...s} fill="#fff" />
    </>
  ),
  freeze: () => (
    <>
      <rect x="1.5" y="1.5" width="13" height="13" {...s} />
      <path d="M1.5 5.5h13M5.5 1.5v13" stroke={BLUE} strokeWidth={1.4} />
    </>
  ),
  print: () => (
    <>
      <rect x="4.5" y="1.5" width="7" height="4" {...s} />
      <rect x="1.5" y="5.5" width="13" height="6" rx="1" {...s} />
      <rect x="4.5" y="9.5" width="7" height="5" {...s} fill="#fff" />
    </>
  ),
  validation: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" {...s} />
      <path d="m4.5 8 2.5 2.5 4.5-5" {...s} stroke={GREEN} strokeWidth={1.4} />
    </>
  ),
  removeDuplicates: () => (
    <>
      <rect x="1.5" y="2.5" width="9" height="11" {...s} />
      <path d="M1.5 6h9M1.5 9.5h9" {...s} />
      <path d="m11 7 3.5 3.5M14.5 7 11 10.5" {...s} stroke={RED} />
    </>
  ),
  conditionalFormat: () => (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" {...s} />
      <path d="M1.5 6h13M1.5 9.5h13M6 2.5v11" {...s} />
      <rect x="6" y="6" width="8.5" height="3.5" fill={RED} opacity="0.45" />
      <rect x="6" y="9.5" width="8.5" height="4" fill={GREEN} opacity="0.4" />
    </>
  ),
  cellStyles: () => (
    <>
      <rect x="1.5" y="3.5" width="13" height="9" {...s} />
      <rect x="1.5" y="3.5" width="6.5" height="4.5" fill="#c6efce" />
      <rect x="8" y="8" width="6.5" height="4.5" fill="#ffc7ce" />
    </>
  ),
  rename: () => <path d="M2.5 13.5 3 11l7.5-7.5 2 2L5 13zM9.5 4.5l2 2" {...s} />,
  trash: () => <path d="M3 4.5h10M6.5 4.5v-2h3v2M4.5 4.5l.7 9.5h5.6l.7-9.5" {...s} />,
  close: () => <path d="m3.5 3.5 9 9M12.5 3.5l-9 9" {...s} strokeWidth={1.3} />,
  check: () => <path d="m2.5 8.5 3.5 3.5 7.5-8" {...s} strokeWidth={1.4} />,
  plus: () => <path d="M8 2.5v11M2.5 8h11" {...s} />,
  minus: () => <path d="M2.5 8h11" {...s} />,
  chevronDown: () => <path d="m4 6 4 4 4-4" {...s} />,
  chevronUp: () => <path d="m4 10 4-4 4 4" {...s} />,
  chevronLeft: () => <path d="M10 4 6 8l4 4" {...s} />,
  chevronRight: () => <path d="m6 4 4 4-4 4" {...s} />,
};

export type IconName = keyof typeof ICONS;

/** Names of all built-in icons (for docs and tests). */
export const iconNames = Object.keys(ICONS);

export function Icon({ name, size = 16, className }: { name: string; size?: number; className?: string }) {
  const draw = ICONS[name];
  if (!draw) return null;
  return (
    <svg
      className={className ? `office-icon ${className}` : "office-icon"}
      width={size}
      height={size}
      viewBox="0 0 16 16"
      aria-hidden="true"
      focusable="false"
    >
      {draw()}
    </svg>
  );
}
