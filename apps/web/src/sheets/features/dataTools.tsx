// Data tools UI: the AutoFilter dropdown, Custom AutoFilter / Top 10, Sort,
// Remove Duplicates, Data Validation (dialog, error alert, in-cell list) and
// Conditional Formatting (highlight rules, data bars, colour scales, Manage
// Rules). The engine does the work (crates/wasm/src/data_tools.rs); SheetsApp
// owns which dialog is open and passes callbacks in, as for layout.tsx.

import { useEffect, useMemo, useRef, useState, type ReactNode, type RefObject } from "react";
import type { SheetsDocument } from "@office/engine";
import type { RibbonCommand, RibbonMenuItem } from "@office/ui";
import { Dialog } from "../Dialog";
import { colName, type Rect } from "../geometry";
import { Select } from "../Select";

// ---------------------------------------------------------------- shared types

/** What `doc.autoFilter` returns. */
export interface FilterInfo {
  range: Rect;
  /** Filtered columns (absolute index). */
  columns: Set<number>;
  sort: { col: number; ascending: boolean } | null;
}

export function readFilter(doc: SheetsDocument, sheet: number): FilterInfo | null {
  const json = doc.autoFilter(sheet);
  if (!json) return null;
  const f = JSON.parse(json) as { range: number[]; columns: Record<string, string>; sort: [number, number] | null };
  const [top, left, bottom, right] = f.range;
  return {
    range: { top, left, bottom, right },
    columns: new Set(Object.keys(f.columns).map(Number)),
    sort: f.sort ? { col: f.sort[0], ascending: f.sort[1] === 1 } : null,
  };
}

/** What `doc.validationAt` / `doc.checkInput` return. */
export interface ValidationRule {
  kind: string;
  operator: string;
  formula1: string;
  formula2: string;
  allowBlank: boolean;
  dropdown: boolean;
  showInput: boolean;
  inputTitle: string;
  inputMessage: string;
  showError: boolean;
  errorStyle: "stop" | "warning" | "information";
  errorTitle: string;
  errorMessage: string;
  ranges: string;
}

export function readValidation(json: string): ValidationRule | null {
  return json ? (JSON.parse(json) as ValidationRule) : null;
}

/** Header text of a column for dialogs: the header cell, or "Column B". */
function columnLabel(doc: SheetsDocument, sheet: number, headerRow: number | null, col: number): string {
  const text = headerRow === null ? "" : doc.displayBlock(sheet, headerRow, col, 1, 1)[0];
  return text || `Column ${colName(col)}`;
}

const button = (label: string, onClick: () => void, opts: { primary?: boolean; disabled?: boolean } = {}) => (
  <button key={label} className={opts.primary ? "dialog-button primary" : "dialog-button"} disabled={opts.disabled} onClick={onClick}>
    {label}
  </button>
);

/**
 * Close a popup when the mouse goes down outside it. Registered on the next
 * tick: the mousedown that opened the popup is still bubbling when React runs
 * the effect, and must not close it again.
 */
function useClickOutside(ref: RefObject<HTMLElement>, onClose: () => void) {
  const latest = useRef(onClose);
  latest.current = onClose;
  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) latest.current();
    };
    const timer = setTimeout(() => window.addEventListener("mousedown", onDown));
    return () => {
      clearTimeout(timer);
      window.removeEventListener("mousedown", onDown);
    };
  }, [ref]);
}

// ---------------------------------------------------------------- ribbon

export interface DataContext {
  filter: FilterInfo | null;
  toggleFilter: () => void;
  clearFilters: () => void;
  reapplyFilter: () => void;
  sort: (ascending: boolean) => void;
  openDialog: (d: DataDialog) => void;
  clearCf: (whole: boolean) => void;
  /** Circle Invalid Data on / off. */
  circles: boolean;
  setCircles: (on: boolean) => void;
}

/** Which data-tools dialog is open (SheetsApp state). */
export type DataDialog =
  | { kind: "sort" }
  | { kind: "dedupe" }
  | { kind: "validation" }
  | { kind: "customFilter"; col: number; preset?: string }
  | { kind: "top10"; col: number }
  | { kind: "cf"; rule: CfDialogKind }
  | { kind: "cfGallery"; gallery: "dataBar" | "colorScale" }
  | { kind: "cfManage" };

export function dataTabCommands(ctx: DataContext): RibbonCommand[] {
  const filtering = !!ctx.filter && ctx.filter.columns.size > 0;
  return [
    { id: "sort-asc", label: "Sort A to Z", title: "Sort by the active column, lowest to highest.", icon: "sortAsc", size: "small", group: "Sort & Filter", onRun: () => ctx.sort(true) },
    { id: "sort-desc", label: "Sort Z to A", title: "Sort by the active column, highest to lowest.", icon: "sortDesc", size: "small", group: "Sort & Filter", onRun: () => ctx.sort(false) },
    { id: "sort-custom", label: "Sort", title: "Sort data by several columns at once.", icon: "sortFilter", size: "large", group: "Sort & Filter", onRun: () => ctx.openDialog({ kind: "sort" }) },
    {
      id: "filter",
      label: "Filter",
      shortcut: "Ctrl+Shift+L",
      title: "Turn on filtering for the selected cells. Then click the arrow in the column header to narrow down the data.",
      icon: "filter",
      size: "large",
      group: "Sort & Filter",
      active: !!ctx.filter,
      onRun: ctx.toggleFilter,
    },
    { id: "filter-clear", label: "Clear", title: "Clear the filter and sort state for the current range of data.", icon: "clear", size: "small", group: "Sort & Filter", disabled: !filtering, onRun: ctx.clearFilters },
    { id: "filter-reapply", label: "Reapply", title: "Reapply the filter to the range after the data changed.", icon: "redo", size: "small", group: "Sort & Filter", disabled: !ctx.filter, onRun: ctx.reapplyFilter },
    {
      id: "remove-duplicates",
      label: "Remove Duplicates",
      title: "Delete duplicate rows from a sheet. You can choose which columns should be checked.",
      icon: "removeDuplicates",
      size: "large",
      group: "Data Tools",
      onRun: () => ctx.openDialog({ kind: "dedupe" }),
    },
    {
      id: "data-validation",
      label: "Data Validation",
      title: "Choose from a list of rules to limit the type of data that can be entered in a cell.",
      icon: "validation",
      size: "large",
      group: "Data Tools",
      onRun: () => ctx.openDialog({ kind: "validation" }),
      menu: [
        { id: "dv-open", label: "Data Validation…", icon: "validation", onRun: () => ctx.openDialog({ kind: "validation" }) },
        { id: "dv-circle", label: "Circle Invalid Data", onRun: () => ctx.setCircles(true), separatorBefore: true },
        { id: "dv-uncircle", label: "Clear Validation Circles", disabled: !ctx.circles, onRun: () => ctx.setCircles(false) },
      ],
    },
  ];
}

/** Items added to Home → Editing → Sort & Filter. */
export function sortFilterMenuItems(ctx: DataContext): RibbonMenuItem[] {
  const filtering = !!ctx.filter && ctx.filter.columns.size > 0;
  return [
    { id: "sort-custom", label: "Custom Sort…", icon: "sortFilter", onRun: () => ctx.openDialog({ kind: "sort" }) },
    { id: "filter-toggle", label: "Filter", shortcut: "Ctrl+Shift+L", icon: "filter", active: !!ctx.filter, onRun: ctx.toggleFilter, separatorBefore: true },
    { id: "filter-clear", label: "Clear", disabled: !filtering, onRun: ctx.clearFilters },
    { id: "filter-reapply", label: "Reapply", disabled: !ctx.filter, onRun: ctx.reapplyFilter },
  ];
}

export function conditionalFormattingCommand(ctx: DataContext): RibbonCommand {
  const cf = (rule: CfDialogKind) => () => ctx.openDialog({ kind: "cf", rule });
  return {
    id: "conditional-formatting",
    label: "Conditional Formatting",
    title: "Easily spot trends and patterns in your data using bars, colors and icons to visually highlight important values.",
    icon: "conditionalFormat",
    size: "large",
    group: "Styles",
    menu: [
      { id: "cf-gt", label: "Greater Than…", onRun: cf("greaterThan") },
      { id: "cf-lt", label: "Less Than…", onRun: cf("lessThan") },
      { id: "cf-between", label: "Between…", onRun: cf("between") },
      { id: "cf-eq", label: "Equal To…", onRun: cf("equal") },
      { id: "cf-text", label: "Text that Contains…", onRun: cf("containsText") },
      { id: "cf-date", label: "A Date Occurring…", onRun: cf("timePeriod") },
      { id: "cf-dup", label: "Duplicate Values…", onRun: cf("duplicates") },
      { id: "cf-top", label: "Top 10 Items…", onRun: cf("top"), separatorBefore: true },
      { id: "cf-top-pct", label: "Top 10%…", onRun: cf("topPercent") },
      { id: "cf-bottom", label: "Bottom 10 Items…", onRun: cf("bottom") },
      { id: "cf-bottom-pct", label: "Bottom 10%…", onRun: cf("bottomPercent") },
      { id: "cf-above", label: "Above Average…", onRun: cf("above") },
      { id: "cf-below", label: "Below Average…", onRun: cf("below") },
      { id: "cf-bars", label: "Data Bars…", onRun: () => ctx.openDialog({ kind: "cfGallery", gallery: "dataBar" }), separatorBefore: true },
      { id: "cf-scales", label: "Color Scales…", onRun: () => ctx.openDialog({ kind: "cfGallery", gallery: "colorScale" }) },
      { id: "cf-new", label: "New Rule (formula)…", onRun: cf("expression"), separatorBefore: true },
      { id: "cf-clear-sel", label: "Clear Rules from Selected Cells", onRun: () => ctx.clearCf(false), separatorBefore: true },
      { id: "cf-clear-all", label: "Clear Rules from Entire Sheet", onRun: () => ctx.clearCf(true) },
      { id: "cf-manage", label: "Manage Rules…", onRun: () => ctx.openDialog({ kind: "cfManage" }), separatorBefore: true },
    ],
  };
}

// ---------------------------------------------------------------- AutoFilter dropdown

/** Above this many distinct values the checklist shows the first ones only. */
const MAX_FILTER_ITEMS = 1000;

interface FilterItem {
  text: string;
  shown: boolean;
  blank: boolean;
}

interface FilterMenuProps {
  doc: SheetsDocument;
  sheet: number;
  filter: FilterInfo;
  col: number;
  /** Screen position of the dropdown's top-left. */
  x: number;
  y: number;
  onClose: () => void;
  /** Called after the engine changed the sheet. */
  onChanged: () => void;
  openDialog: (d: DataDialog) => void;
}

/** Excel's filter dropdown: sort, clear, Text/Number Filters, search and the value checklist. */
export function FilterMenu({ doc, sheet, filter, col, x, y, onClose, onChanged, openDialog }: FilterMenuProps) {
  const items = useMemo(() => JSON.parse(doc.filterItems(sheet, col)) as FilterItem[], [doc, sheet, col]);
  const [checked, setChecked] = useState(() => items.map((i) => i.shown));
  const [search, setSearch] = useState("");
  const ref = useRef<HTMLDivElement>(null);
  const header = columnLabel(doc, sheet, filter.range.top, col);
  const filtered = filter.columns.has(col);
  const numeric = items.length > 0 && items.every((i) => i.blank || /^[-+]?[₹$]?[\d,]*\.?\d+%?$/.test(i.text.trim()));

  useClickOutside(ref, onClose);

  const q = search.trim().toLowerCase();
  const visible = items.map((item, i) => ({ item, i })).filter(({ item }) => !q || (!item.blank && item.text.toLowerCase().includes(q)));
  const shownItems = visible.slice(0, MAX_FILTER_ITEMS);
  const allChecked = visible.every(({ i }) => checked[i]);
  const noneChecked = visible.every(({ i }) => !checked[i]);

  function setAll(on: boolean) {
    setChecked((c) => {
      const next = [...c];
      for (const { i } of visible) next[i] = on;
      return next;
    });
  }

  function run(action: () => void) {
    action();
    onChanged();
    onClose();
  }

  function ok() {
    // Searching filters to the matching values that are ticked, like Excel.
    const pick = q ? visible : items.map((item, i) => ({ item, i }));
    const everything = !q && checked.every(Boolean);
    run(() => {
      if (everything) doc.clearColumnFilter(sheet, col);
      else {
        const values = pick.filter(({ item, i }) => checked[i] && !item.blank).map(({ item }) => item.text);
        const blanks = pick.some(({ item, i }) => item.blank && checked[i]);
        doc.filterByValues(sheet, col, values, blanks);
      }
    });
  }

  const sortActive = filter.sort?.col === col ? filter.sort.ascending : null;
  return (
    <div
      ref={ref}
      className="filter-menu"
      role="dialog"
      aria-label={`Filter ${header}`}
      style={{ left: Math.min(x, window.innerWidth - 270), top: Math.min(y, window.innerHeight - 440) }}
      onKeyDown={(e) => {
        if (e.key === "Escape") {
          e.stopPropagation();
          onClose();
        }
      }}
    >
      <button className={sortActive === true ? "filter-menu-item active" : "filter-menu-item"} onClick={() => run(() => doc.filterSort(sheet, col, true))}>
        {numeric ? "Sort Smallest to Largest" : "Sort A to Z"}
      </button>
      <button className={sortActive === false ? "filter-menu-item active" : "filter-menu-item"} onClick={() => run(() => doc.filterSort(sheet, col, false))}>
        {numeric ? "Sort Largest to Smallest" : "Sort Z to A"}
      </button>
      <hr />
      <button className="filter-menu-item" disabled={!filtered} onClick={() => run(() => doc.clearColumnFilter(sheet, col))}>
        Clear Filter From “{header}”
      </button>
      <div className="filter-menu-row">
        <button
          className="filter-menu-item"
          onClick={() => {
            onClose();
            openDialog({ kind: "customFilter", col, preset: numeric ? "greaterThan" : "contains" });
          }}
        >
          {numeric ? "Number Filters…" : "Text Filters…"}
        </button>
        {numeric && (
          <>
            <button
              className="filter-menu-chip"
              onClick={() => {
                onClose();
                openDialog({ kind: "top10", col });
              }}
            >
              Top 10…
            </button>
            <button className="filter-menu-chip" onClick={() => run(() => doc.filterAverage(sheet, col, true))}>
              Above Avg
            </button>
            <button className="filter-menu-chip" onClick={() => run(() => doc.filterAverage(sheet, col, false))}>
              Below Avg
            </button>
          </>
        )}
      </div>
      <input
        className="filter-menu-search"
        placeholder="Search"
        aria-label="Search values"
        value={search}
        autoFocus
        onChange={(e) => setSearch(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && !noneChecked && ok()}
      />
      <div className="filter-menu-list" role="list">
        <label className="filter-menu-check">
          <input
            type="checkbox"
            checked={allChecked && visible.length > 0}
            ref={(el) => {
              if (el) el.indeterminate = !allChecked && !noneChecked;
            }}
            onChange={() => setAll(!allChecked)}
          />
          {q ? "(Select All Search Results)" : "(Select All)"}
        </label>
        {shownItems.map(({ item, i }) => (
          <label key={i} className="filter-menu-check">
            <input type="checkbox" checked={checked[i]} onChange={() => setChecked((c) => c.map((v, j) => (j === i ? !v : v)))} />
            {item.blank ? "(Blanks)" : item.text}
          </label>
        ))}
        {visible.length > shownItems.length && <p className="filter-menu-note">Not all items showing. Search to find more.</p>}
      </div>
      <div className="filter-menu-footer">
        {button("OK", ok, { primary: true, disabled: noneChecked })}
        {button("Cancel", onClose)}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------- Custom AutoFilter, Top 10

/** Custom AutoFilter operators; begins/ends/contains are stored as wildcards, as Excel does. */
const FILTER_OPS: { id: string; label: string; op: string; wrap: (v: string) => string }[] = [
  { id: "equal", label: "equals", op: "equal", wrap: (v) => v },
  { id: "notEqual", label: "does not equal", op: "notEqual", wrap: (v) => v },
  { id: "greaterThan", label: "is greater than", op: "greaterThan", wrap: (v) => v },
  { id: "greaterThanOrEqual", label: "is greater than or equal to", op: "greaterThanOrEqual", wrap: (v) => v },
  { id: "lessThan", label: "is less than", op: "lessThan", wrap: (v) => v },
  { id: "lessThanOrEqual", label: "is less than or equal to", op: "lessThanOrEqual", wrap: (v) => v },
  { id: "beginsWith", label: "begins with", op: "equal", wrap: (v) => `${v}*` },
  { id: "notBeginsWith", label: "does not begin with", op: "notEqual", wrap: (v) => `${v}*` },
  { id: "endsWith", label: "ends with", op: "equal", wrap: (v) => `*${v}` },
  { id: "notEndsWith", label: "does not end with", op: "notEqual", wrap: (v) => `*${v}` },
  { id: "contains", label: "contains", op: "equal", wrap: (v) => `*${v}*` },
  { id: "notContains", label: "does not contain", op: "notEqual", wrap: (v) => `*${v}*` },
];

export function CustomFilterDialog(props: { doc: SheetsDocument; sheet: number; filter: FilterInfo; col: number; preset?: string; onDone: () => void; onClose: () => void }) {
  const { doc, sheet, filter, col } = props;
  const [ops, setOps] = useState([props.preset ?? "equal", "equal"]);
  const [values, setValues] = useState(["", ""]);
  const [and, setAnd] = useState(true);
  const header = columnLabel(doc, sheet, filter.range.top, col);
  function ok() {
    const pick = (i: number) => FILTER_OPS.find((o) => o.id === ops[i])!;
    const used = [0, 1].filter((i) => values[i].trim() !== "");
    doc.filterByConditions(
      sheet,
      col,
      and,
      used.map((i) => pick(i).op),
      used.map((i) => pick(i).wrap(values[i].trim())),
    );
    props.onDone();
    props.onClose();
  }
  const row = (i: number) => (
    <div className="dialog-row" key={i}>
      <Select aria-label={`Condition ${i + 1}`} value={ops[i]} onChange={(e) => setOps((o) => o.map((v, j) => (j === i ? e.target.value : v)))}>
        {FILTER_OPS.map((o) => (
          <option key={o.id} value={o.id}>
            {o.label}
          </option>
        ))}
      </Select>
      <input
        aria-label={`Value ${i + 1}`}
        autoFocus={i === 0}
        value={values[i]}
        onChange={(e) => setValues((v) => v.map((x, j) => (j === i ? e.target.value : x)))}
        onKeyDown={(e) => e.key === "Enter" && ok()}
      />
    </div>
  );
  return (
    <Dialog title="Custom AutoFilter" onClose={props.onClose} width={460} footer={<>{button("OK", ok, { primary: true })}{button("Cancel", props.onClose)}</>}>
      <p className="dialog-label">Show rows where: {header}</p>
      {row(0)}
      <div className="dialog-row radios">
        <label>
          <input type="radio" checked={and} onChange={() => setAnd(true)} /> And
        </label>
        <label>
          <input type="radio" checked={!and} onChange={() => setAnd(false)} /> Or
        </label>
      </div>
      {row(1)}
      <p className="dialog-hint">Use ? to represent any single character. Use * to represent any series of characters.</p>
    </Dialog>
  );
}

export function Top10Dialog(props: { doc: SheetsDocument; sheet: number; col: number; onDone: () => void; onClose: () => void }) {
  const [top, setTop] = useState(true);
  const [value, setValue] = useState("10");
  const [percent, setPercent] = useState(false);
  const n = Number(value);
  const valid = n > 0 && n <= (percent ? 100 : 500);
  function ok() {
    if (!valid) return;
    props.doc.filterTop10(props.sheet, props.col, top, percent, n);
    props.onDone();
    props.onClose();
  }
  return (
    <Dialog title="Top 10 AutoFilter" onClose={props.onClose} width={380} footer={<>{button("OK", ok, { primary: true, disabled: !valid })}{button("Cancel", props.onClose)}</>}>
      <div className="dialog-row">
        <Select aria-label="Top or bottom" value={top ? "top" : "bottom"} onChange={(e) => setTop(e.target.value === "top")}>
          <option value="top">Top</option>
          <option value="bottom">Bottom</option>
        </Select>
        <input aria-label="How many" inputMode="numeric" value={value} autoFocus onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} style={{ width: 70 }} />
        <Select aria-label="Items or percent" value={percent ? "percent" : "items"} onChange={(e) => setPercent(e.target.value === "percent")}>
          <option value="items">Items</option>
          <option value="percent">Percent</option>
        </Select>
      </div>
    </Dialog>
  );
}

// ---------------------------------------------------------------- Sort

interface SortLevel {
  col: number;
  ascending: boolean;
}

/** The area Sort and Remove Duplicates work on: the selection, or the table around a single cell. */
export function tableArea(doc: SheetsDocument, sheet: number, sel: Rect): Rect {
  if (sel.top !== sel.bottom || sel.left !== sel.right) {
    const [rows, cols] = doc.extent(sheet);
    return { ...sel, bottom: Math.min(sel.bottom, Math.max(rows - 1, sel.top)), right: Math.min(sel.right, Math.max(cols - 1, sel.left)) };
  }
  const [top, left, bottom, right] = doc.currentRegion(sheet, sel.top, sel.left);
  return { top, left, bottom, right };
}

/** Excel guesses a header row when the first row is text above numbers or other kinds of value. */
export function guessHeader(doc: SheetsDocument, sheet: number, area: Rect): boolean {
  if (area.bottom <= area.top) return false;
  const width = area.right - area.left + 1;
  const first = doc.kindBlock(sheet, area.top, area.left, 1, width);
  const second = doc.kindBlock(sheet, area.top + 1, area.left, 1, width);
  return [...first].every((k) => k === 2 || k === 0) && [...second].some((k, i) => k !== first[i] && k !== 0);
}

export function SortDialog(props: { doc: SheetsDocument; sheet: number; sel: Rect; activeCol: number; onDone: (area: Rect) => void; onClose: () => void }) {
  const { doc, sheet } = props;
  const area = useMemo(() => tableArea(doc, sheet, props.sel), [doc, sheet, props.sel]);
  const [header, setHeader] = useState(() => guessHeader(doc, sheet, area));
  const startCol = props.activeCol >= area.left && props.activeCol <= area.right ? props.activeCol : area.left;
  const [levels, setLevels] = useState<SortLevel[]>([{ col: startCol, ascending: true }]);
  const [current, setCurrent] = useState(0);
  const columns = Array.from({ length: area.right - area.left + 1 }, (_, i) => area.left + i);
  const label = (c: number) => (header ? columnLabel(doc, sheet, area.top, c) : `Column ${colName(c)}`);

  const update = (i: number, patch: Partial<SortLevel>) => setLevels((ls) => ls.map((l, j) => (j === i ? { ...l, ...patch } : l)));
  function move(d: number) {
    const j = current + d;
    if (j < 0 || j >= levels.length) return;
    setLevels((ls) => {
      const next = [...ls];
      [next[current], next[j]] = [next[j], next[current]];
      return next;
    });
    setCurrent(j);
  }
  function ok() {
    const top = area.top + (header ? 1 : 0);
    if (top >= area.bottom) {
      props.onClose();
      return;
    }
    if (!doc.canSort(sheet, top, area.left, area.bottom, area.right)) {
      alert("To do this, all the merged cells need to be the same size.");
      return;
    }
    doc.sortBy(sheet, top, area.left, area.bottom, area.right, Uint32Array.from(levels.map((l) => l.col)), Uint8Array.from(levels.map((l) => (l.ascending ? 1 : 0))));
    props.onDone(area);
    props.onClose();
  }
  return (
    <Dialog title="Sort" onClose={props.onClose} width={560} footer={<>{button("OK", ok, { primary: true })}{button("Cancel", props.onClose)}</>}>
      <div className="dialog-row toolbar">
        <button className="dialog-button" disabled={levels.length >= 64} onClick={() => {
          setLevels((ls) => [...ls, { col: area.left, ascending: true }]);
          setCurrent(levels.length);
        }}>
          + Add Level
        </button>
        <button className="dialog-button" disabled={levels.length <= 1} onClick={() => {
          setLevels((ls) => ls.filter((_, j) => j !== current));
          setCurrent((c) => Math.max(0, c - 1));
        }}>
          Delete Level
        </button>
        <button className="dialog-button" onClick={() => {
          setLevels((ls) => [...ls.slice(0, current + 1), { ...ls[current] }, ...ls.slice(current + 1)]);
          setCurrent(current + 1);
        }}>
          Copy Level
        </button>
        <button className="dialog-button icon" aria-label="Move up" disabled={current === 0} onClick={() => move(-1)}>▲</button>
        <button className="dialog-button icon" aria-label="Move down" disabled={current === levels.length - 1} onClick={() => move(1)}>▼</button>
        <label className="dialog-check push">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> My data has headers
        </label>
      </div>
      <table className="dialog-table">
        <thead>
          <tr>
            <th />
            <th>Column</th>
            <th>Sort On</th>
            <th>Order</th>
          </tr>
        </thead>
        <tbody>
          {levels.map((l, i) => (
            <tr key={i} className={i === current ? "current" : undefined} onMouseDown={() => setCurrent(i)}>
              <td>{i === 0 ? "Sort by" : "Then by"}</td>
              <td>
                <Select aria-label={`Level ${i + 1} column`} value={l.col} onChange={(e) => update(i, { col: Number(e.target.value) })}>
                  {columns.map((c) => (
                    <option key={c} value={c}>
                      {label(c)}
                    </option>
                  ))}
                </Select>
              </td>
              <td>Cell Values</td>
              <td>
                <Select aria-label={`Level ${i + 1} order`} value={l.ascending ? "asc" : "desc"} onChange={(e) => update(i, { ascending: e.target.value === "asc" })}>
                  <option value="asc">A to Z / Smallest to Largest</option>
                  <option value="desc">Z to A / Largest to Smallest</option>
                </Select>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </Dialog>
  );
}

// ---------------------------------------------------------------- Remove Duplicates

export function RemoveDuplicatesDialog(props: { doc: SheetsDocument; sheet: number; sel: Rect; onDone: (area: Rect, message: string) => void; onClose: () => void }) {
  const { doc, sheet } = props;
  const area = useMemo(() => tableArea(doc, sheet, props.sel), [doc, sheet, props.sel]);
  const [header, setHeader] = useState(() => guessHeader(doc, sheet, area));
  const columns = Array.from({ length: area.right - area.left + 1 }, (_, i) => area.left + i);
  const [picked, setPicked] = useState(() => new Set(columns));
  function ok() {
    const [removed, remaining] = doc.removeDuplicates(sheet, area.top, area.left, area.bottom, area.right, Uint32Array.from([...picked].sort((a, b) => a - b)), header);
    const message =
      removed === 0
        ? "No duplicate values found."
        : `${removed} duplicate value${removed === 1 ? "" : "s"} found and removed; ${remaining} unique value${remaining === 1 ? "" : "s"} remain.`;
    props.onDone(area, message);
    props.onClose();
  }
  return (
    <Dialog title="Remove Duplicates" onClose={props.onClose} width={420} footer={<>{button("OK", ok, { primary: true, disabled: picked.size === 0 })}{button("Cancel", props.onClose)}</>}>
      <p className="dialog-label">To delete duplicate values, select one or more columns that contain duplicates.</p>
      <div className="dialog-row toolbar">
        <button className="dialog-button" onClick={() => setPicked(new Set(columns))}>Select All</button>
        <button className="dialog-button" onClick={() => setPicked(new Set())}>Unselect All</button>
        <label className="dialog-check push">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> My data has headers
        </label>
      </div>
      <div className="dialog-checklist" role="list" aria-label="Columns">
        {columns.map((c) => (
          <label key={c} className="dialog-check">
            <input
              type="checkbox"
              checked={picked.has(c)}
              onChange={() =>
                setPicked((p) => {
                  const next = new Set(p);
                  if (next.has(c)) next.delete(c);
                  else next.add(c);
                  return next;
                })
              }
            />
            {header ? columnLabel(doc, sheet, area.top, c) : `Column ${colName(c)}`}
          </label>
        ))}
      </div>
    </Dialog>
  );
}

/** A plain message box with OK. */
export function MessageDialog({ title, message, onClose }: { title: string; message: string; onClose: () => void }) {
  return (
    <Dialog title={title} onClose={onClose} width={380} footer={button("OK", onClose, { primary: true })}>
      <p className="dialog-message">{message}</p>
    </Dialog>
  );
}

// ---------------------------------------------------------------- Data Validation

const VALIDATION_KINDS = [
  ["none", "Any value"],
  ["whole", "Whole number"],
  ["decimal", "Decimal"],
  ["list", "List"],
  ["date", "Date"],
  ["time", "Time"],
  ["textLength", "Text length"],
  ["custom", "Custom"],
] as const;

const VALIDATION_OPS = [
  ["between", "between"],
  ["notBetween", "not between"],
  ["equal", "equal to"],
  ["notEqual", "not equal to"],
  ["greaterThan", "greater than"],
  ["lessThan", "less than"],
  ["greaterThanOrEqual", "greater than or equal to"],
  ["lessThanOrEqual", "less than or equal to"],
] as const;

const EMPTY_RULE: ValidationRule = {
  kind: "none",
  operator: "between",
  formula1: "",
  formula2: "",
  allowBlank: true,
  dropdown: true,
  showInput: true,
  inputTitle: "",
  inputMessage: "",
  showError: true,
  errorStyle: "stop",
  errorTitle: "",
  errorMessage: "",
  ranges: "",
};

/** List sources typed as `"a,b"` show in the dialog without the quotes, like Excel. */
function sourceForDialog(f: string): string {
  const m = /^"(.*)"$/.exec(f);
  return m ? m[1].replace(/""/g, '"') : f && !/^[-\d.]/.test(f) && !f.includes(",") ? `=${f}` : f;
}

function sourceForEngine(kind: string, s: string): string {
  const t = s.trim();
  if (kind !== "list" || t.startsWith("=") || t === "") return t;
  return `"${t.replace(/"/g, '""')}"`;
}

export function ValidationDialog(props: { doc: SheetsDocument; sheet: number; sel: Rect; active: { row: number; col: number }; onDone: () => void; onClose: () => void }) {
  const { doc, sheet, sel } = props;
  const initial = readValidation(doc.validationAt(sheet, props.active.row, props.active.col)) ?? EMPTY_RULE;
  const [rule, setRule] = useState<ValidationRule>({
    ...initial,
    formula1: initial.kind === "list" ? sourceForDialog(initial.formula1) : initial.formula1 && `=${initial.formula1}`.replace(/^=(-?[\d.]+)$/, "$1"),
    formula2: initial.formula2 && `=${initial.formula2}`.replace(/^=(-?[\d.]+)$/, "$1"),
  });
  const [tab, setTab] = useState<"settings" | "input" | "error">("settings");
  const set = (patch: Partial<ValidationRule>) => setRule((r) => ({ ...r, ...patch }));
  const two = rule.operator === "between" || rule.operator === "notBetween";
  const ranged = ["whole", "decimal", "date", "time", "textLength"].includes(rule.kind);
  const fieldName = rule.kind === "date" ? "date" : rule.kind === "time" ? "time" : rule.kind === "textLength" ? "length" : "value";
  const cap = (s: string) => s[0].toUpperCase() + s.slice(1);

  function ok() {
    doc.setValidation(
      sheet,
      sel.top,
      sel.left,
      sel.bottom,
      sel.right,
      rule.kind,
      rule.operator,
      sourceForEngine(rule.kind, rule.formula1),
      rule.formula2.trim(),
      rule.allowBlank,
      rule.dropdown,
      rule.showInput,
      rule.inputTitle,
      rule.inputMessage,
      rule.showError,
      rule.errorStyle,
      rule.errorTitle,
      rule.errorMessage,
    );
    props.onDone();
    props.onClose();
  }
  function clearAll() {
    doc.clearValidation(sheet, sel.top, sel.left, sel.bottom, sel.right);
    props.onDone();
    props.onClose();
  }

  const tabs: [typeof tab, string][] = [
    ["settings", "Settings"],
    ["input", "Input Message"],
    ["error", "Error Alert"],
  ];
  let body: ReactNode;
  if (tab === "settings") {
    body = (
      <>
        <p className="dialog-label">Validation criteria</p>
        <label className="dialog-field">
          <span>Allow:</span>
          <Select aria-label="Allow" value={rule.kind} onChange={(e) => set({ kind: e.target.value })}>
            {VALIDATION_KINDS.map(([id, label]) => (
              <option key={id} value={id}>
                {label}
              </option>
            ))}
          </Select>
        </label>
        {rule.kind !== "none" && (
          <label className="dialog-check">
            <input type="checkbox" checked={rule.allowBlank} onChange={(e) => set({ allowBlank: e.target.checked })} /> Ignore blank
          </label>
        )}
        {rule.kind === "list" && (
          <label className="dialog-check">
            <input type="checkbox" checked={rule.dropdown} onChange={(e) => set({ dropdown: e.target.checked })} /> In-cell dropdown
          </label>
        )}
        {ranged && (
          <label className="dialog-field">
            <span>Data:</span>
            <Select aria-label="Data" value={rule.operator} onChange={(e) => set({ operator: e.target.value })}>
              {VALIDATION_OPS.map(([id, label]) => (
                <option key={id} value={id}>
                  {label}
                </option>
              ))}
            </Select>
          </label>
        )}
        {ranged && (
          <label className="dialog-field">
            <span>{two ? `Minimum ${fieldName}:` : `${cap(fieldName)}:`}</span>
            <input aria-label={two ? "Minimum" : "Value"} value={rule.formula1} onChange={(e) => set({ formula1: e.target.value })} />
          </label>
        )}
        {ranged && two && (
          <label className="dialog-field">
            <span>Maximum {fieldName}:</span>
            <input aria-label="Maximum" value={rule.formula2} onChange={(e) => set({ formula2: e.target.value })} />
          </label>
        )}
        {rule.kind === "list" && (
          <label className="dialog-field">
            <span>Source:</span>
            <input aria-label="Source" placeholder="Yes,No or =$A$1:$A$5" value={rule.formula1} onChange={(e) => set({ formula1: e.target.value })} />
          </label>
        )}
        {rule.kind === "custom" && (
          <label className="dialog-field">
            <span>Formula:</span>
            <input aria-label="Formula" placeholder="=A1>0" value={rule.formula1} onChange={(e) => set({ formula1: e.target.value })} />
          </label>
        )}
      </>
    );
  } else if (tab === "input") {
    body = (
      <>
        <label className="dialog-check">
          <input type="checkbox" checked={rule.showInput} onChange={(e) => set({ showInput: e.target.checked })} /> Show input message when cell is selected
        </label>
        <p className="dialog-label">When cell is selected, show this input message:</p>
        <label className="dialog-field">
          <span>Title:</span>
          <input aria-label="Input title" maxLength={32} value={rule.inputTitle} onChange={(e) => set({ inputTitle: e.target.value })} />
        </label>
        <label className="dialog-field">
          <span>Input message:</span>
          <textarea aria-label="Input message" maxLength={255} rows={4} value={rule.inputMessage} onChange={(e) => set({ inputMessage: e.target.value })} />
        </label>
      </>
    );
  } else {
    body = (
      <>
        <label className="dialog-check">
          <input type="checkbox" checked={rule.showError} onChange={(e) => set({ showError: e.target.checked })} /> Show error alert after invalid data is entered
        </label>
        <p className="dialog-label">When user enters invalid data, show this error alert:</p>
        <label className="dialog-field">
          <span>Style:</span>
          <Select aria-label="Style" value={rule.errorStyle} onChange={(e) => set({ errorStyle: e.target.value as ValidationRule["errorStyle"] })}>
            <option value="stop">Stop</option>
            <option value="warning">Warning</option>
            <option value="information">Information</option>
          </Select>
        </label>
        <label className="dialog-field">
          <span>Title:</span>
          <input aria-label="Error title" maxLength={32} value={rule.errorTitle} onChange={(e) => set({ errorTitle: e.target.value })} />
        </label>
        <label className="dialog-field">
          <span>Error message:</span>
          <textarea aria-label="Error message" maxLength={255} rows={4} value={rule.errorMessage} onChange={(e) => set({ errorMessage: e.target.value })} />
        </label>
      </>
    );
  }
  return (
    <Dialog
      title="Data Validation"
      onClose={props.onClose}
      width={440}
      footer={
        <>
          <button className="dialog-button push-left" onClick={clearAll}>
            Clear All
          </button>
          {button("OK", ok, { primary: true })}
          {button("Cancel", props.onClose)}
        </>
      }
    >
      <div className="dialog-tabs" role="tablist">
        {tabs.map(([id, label]) => (
          <button key={id} role="tab" aria-selected={tab === id} className={tab === id ? "dialog-tab active" : "dialog-tab"} onClick={() => setTab(id)}>
            {label}
          </button>
        ))}
      </div>
      <div className="dialog-tab-body">{body}</div>
    </Dialog>
  );
}

/** What happens after an entry breaks a validation rule. */
export type AlertChoice = "retry" | "cancel" | "accept";

/** The error alert shown when a typed entry breaks a rule (Stop / Warning / Information). */
export function ValidationAlert({ rule, onChoice }: { rule: ValidationRule; onChoice: (c: AlertChoice) => void }) {
  const message = rule.errorMessage || "This value doesn't match the data validation restrictions defined for this cell.";
  const title = rule.errorTitle || "Sheets";
  const icon = rule.errorStyle === "stop" ? "✖" : rule.errorStyle === "warning" ? "⚠" : "ℹ";
  const buttons =
    rule.errorStyle === "stop"
      ? [button("Retry", () => onChoice("retry"), { primary: true }), button("Cancel", () => onChoice("cancel"))]
      : rule.errorStyle === "warning"
        ? [button("Yes", () => onChoice("accept"), { primary: true }), button("No", () => onChoice("retry")), button("Cancel", () => onChoice("cancel"))]
        : [button("OK", () => onChoice("accept"), { primary: true }), button("Cancel", () => onChoice("cancel"))];
  return (
    <Dialog title={title} onClose={() => onChoice("cancel")} width={400} className="validation-alert" footer={<>{buttons}</>}>
      <div className={`alert-body ${rule.errorStyle}`}>
        <span className="alert-icon" aria-hidden>
          {icon}
        </span>
        <p>
          {message}
          {rule.errorStyle === "warning" && (
            <>
              <br />
              <br />
              Continue?
            </>
          )}
        </p>
      </div>
    </Dialog>
  );
}

/** The in-cell dropdown of a List rule. */
export function ListDropdown({ items, x, y, minWidth, onPick, onClose }: { items: string[]; x: number; y: number; minWidth: number; onPick: (s: string) => void; onClose: () => void }) {
  const ref = useRef<HTMLUListElement>(null);
  const [index, setIndex] = useState(0);
  useEffect(() => ref.current?.focus(), []);
  useClickOutside(ref, onClose);
  return (
    <ul
      ref={ref}
      className="list-dropdown"
      role="listbox"
      tabIndex={-1}
      style={{ left: x, top: y, minWidth }}
      onKeyDown={(e) => {
        if (e.key === "ArrowDown") setIndex((i) => Math.min(items.length - 1, i + 1));
        else if (e.key === "ArrowUp") setIndex((i) => Math.max(0, i - 1));
        else if (e.key === "Enter") onPick(items[index]);
        else if (e.key === "Escape") onClose();
        else return;
        e.preventDefault();
        e.stopPropagation();
      }}
    >
      {items.map((item, i) => (
        <li key={i} role="option" aria-selected={i === index} className={i === index ? "active" : undefined} onMouseEnter={() => setIndex(i)} onMouseDown={(e) => {
          e.preventDefault();
          onPick(item);
        }}>
          {item}
        </li>
      ))}
    </ul>
  );
}

// ---------------------------------------------------------------- Conditional Formatting

export type CfDialogKind =
  | "greaterThan"
  | "lessThan"
  | "between"
  | "equal"
  | "containsText"
  | "timePeriod"
  | "duplicates"
  | "top"
  | "topPercent"
  | "bottom"
  | "bottomPercent"
  | "above"
  | "below"
  | "expression";

const CF_FORMATS = [
  ["lightRedDarkRed", "Light Red Fill with Dark Red Text", "#FFC7CE", "#9C0006"],
  ["yellowDarkYellow", "Yellow Fill with Dark Yellow Text", "#FFEB9C", "#9C5700"],
  ["greenDarkGreen", "Green Fill with Dark Green Text", "#C6EFCE", "#006100"],
  ["lightRed", "Light Red Fill", "#FFC7CE", ""],
  ["redText", "Red Text", "", "#9C0006"],
  ["fill=#BDD7EE;color=#1F4E78;b=1", "Blue Fill with Bold Dark Blue Text", "#BDD7EE", "#1F4E78"],
] as const;

const TIME_PERIODS = [
  ["yesterday", "Yesterday"],
  ["today", "Today"],
  ["tomorrow", "Tomorrow"],
  ["last7Days", "In the last 7 days"],
  ["lastWeek", "Last week"],
  ["thisWeek", "This week"],
  ["nextWeek", "Next week"],
  ["lastMonth", "Last month"],
  ["thisMonth", "This month"],
  ["nextMonth", "Next month"],
] as const;

const CF_TEXT: Record<CfDialogKind, { title: string; prompt: string }> = {
  greaterThan: { title: "Greater Than", prompt: "Format cells that are GREATER THAN:" },
  lessThan: { title: "Less Than", prompt: "Format cells that are LESS THAN:" },
  between: { title: "Between", prompt: "Format cells that are BETWEEN:" },
  equal: { title: "Equal To", prompt: "Format cells that are EQUAL TO:" },
  containsText: { title: "Text That Contains", prompt: "Format cells that contain the text:" },
  timePeriod: { title: "A Date Occurring", prompt: "Format cells that contain a date occurring:" },
  duplicates: { title: "Duplicate Values", prompt: "Format cells that contain:" },
  top: { title: "Top 10 Items", prompt: "Format cells that rank in the TOP:" },
  topPercent: { title: "Top 10%", prompt: "Format cells that rank in the TOP:" },
  bottom: { title: "Bottom 10 Items", prompt: "Format cells that rank in the BOTTOM:" },
  bottomPercent: { title: "Bottom 10%", prompt: "Format cells that rank in the BOTTOM:" },
  above: { title: "Above Average", prompt: "Format cells that are ABOVE AVERAGE:" },
  below: { title: "Below Average", prompt: "Format cells that are BELOW AVERAGE:" },
  expression: { title: "New Formatting Rule", prompt: "Format values where this formula is true:" },
};

/** A starting value for Greater Than etc.: Excel suggests a value from the selection. */
function suggestedValue(doc: SheetsDocument, sheet: number, sel: Rect): string {
  const [count, numbers, sum] = doc.rangeStats(sheet, sel.top, sel.left, Math.min(sel.bottom - sel.top + 1, 100_000), sel.right - sel.left + 1);
  return count > 0 && numbers > 0 ? String(Math.round((sum / numbers) * 100) / 100) : "";
}

export function CfRuleDialog(props: { doc: SheetsDocument; sheet: number; sel: Rect; rule: CfDialogKind; onDone: () => void; onClose: () => void }) {
  const { doc, sheet, sel, rule } = props;
  const text = CF_TEXT[rule];
  const ranked = rule === "top" || rule === "topPercent" || rule === "bottom" || rule === "bottomPercent";
  const [v1, setV1] = useState(() => (ranked ? "10" : rule === "expression" ? "=" : ["greaterThan", "lessThan", "between", "equal"].includes(rule) ? suggestedValue(doc, sheet, sel) : ""));
  const [v2, setV2] = useState("");
  const [choice, setChoice] = useState(rule === "timePeriod" ? "yesterday" : rule === "duplicates" ? "duplicates" : "");
  const [format, setFormat] = useState<string>(CF_FORMATS[0][0]);
  const [error, setError] = useState("");

  function spec(): [string, string[]] {
    switch (rule) {
      case "greaterThan":
      case "lessThan":
      case "equal":
        return ["cellIs", [rule, v1, ""]];
      case "between":
        return ["cellIs", ["between", v1, v2]];
      case "containsText":
        return ["text", ["containsText", v1]];
      case "timePeriod":
        return ["timePeriod", [choice]];
      case "duplicates":
        return [choice === "unique" ? "unique" : "duplicates", []];
      case "top":
      case "topPercent":
      case "bottom":
      case "bottomPercent":
        return ["top10", [rule.startsWith("top") ? "top" : "bottom", rule.endsWith("Percent") ? "percent" : "items", v1]];
      case "above":
      case "below":
        return ["average", [rule]];
      case "expression":
        return ["expression", [v1]];
    }
  }
  function ok() {
    const [kind, args] = spec();
    if (!doc.addCfRule(sheet, sel.top, sel.left, sel.bottom, sel.right, kind, args, format)) {
      setError(ranked ? "Enter a whole number between 1 and 1000." : "Enter a valid value.");
      return;
    }
    props.onDone();
    props.onClose();
  }

  const input = (value: string, set: (v: string) => void, label: string, focus = false) => (
    <input aria-label={label} autoFocus={focus} value={value} onChange={(e) => set(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} />
  );
  let field: ReactNode = null;
  if (rule === "between") {
    field = (
      <div className="dialog-row">
        {input(v1, setV1, "From", true)} <span>and</span> {input(v2, setV2, "To")}
      </div>
    );
  } else if (rule === "timePeriod") {
    field = (
      <Select aria-label="Date occurring" value={choice} onChange={(e) => setChoice(e.target.value)} autoFocus>
        {TIME_PERIODS.map(([id, label]) => (
          <option key={id} value={id}>
            {label}
          </option>
        ))}
      </Select>
    );
  } else if (rule === "duplicates") {
    field = (
      <Select aria-label="Duplicate or unique" value={choice} onChange={(e) => setChoice(e.target.value)} autoFocus>
        <option value="duplicates">Duplicate</option>
        <option value="unique">Unique</option>
      </Select>
    );
  } else if (rule !== "above" && rule !== "below") {
    field = (
      <div className="dialog-row">
        {input(v1, setV1, text.prompt, true)}
        {ranked && <span>{rule.endsWith("Percent") ? "%" : "items"}</span>}
      </div>
    );
  }
  const preview = CF_FORMATS.find((f) => f[0] === format)!;
  return (
    <Dialog title={text.title} onClose={props.onClose} width={480} footer={<>{button("OK", ok, { primary: true })}{button("Cancel", props.onClose)}</>}>
      <p className="dialog-label">{text.prompt}</p>
      {field}
      <div className="dialog-row">
        <span>{rule === "duplicates" ? "values with" : "with"}</span>
        <Select aria-label="Format" value={format} onChange={(e) => setFormat(e.target.value)}>
          {CF_FORMATS.map(([id, label]) => (
            <option key={id} value={id}>
              {label}
            </option>
          ))}
        </Select>
        <span className="cf-preview" style={{ background: preview[2] || "#fff", color: preview[3] || "#000" }}>
          AaBbCc
        </span>
      </div>
      {error && <p className="dialog-error">{error}</p>}
    </Dialog>
  );
}

const DATA_BARS = [
  ["#638EC6", "Blue"],
  ["#63BE7B", "Green"],
  ["#FF555A", "Red"],
  ["#FFB628", "Orange"],
  ["#008AEF", "Light Blue"],
  ["#D6007B", "Purple"],
] as const;

/** Lowest-value colour first, as `addCfRule` takes them. */
const COLOR_SCALES: [string[], string][] = [
  [["#F8696B", "#FFEB84", "#63BE7B"], "Green - Yellow - Red (high is green)"],
  [["#63BE7B", "#FFEB84", "#F8696B"], "Red - Yellow - Green (high is red)"],
  [["#F8696B", "#FCFCFF", "#63BE7B"], "Green - White - Red"],
  [["#63BE7B", "#FCFCFF", "#F8696B"], "Red - White - Green"],
  [["#F8696B", "#FCFCFF", "#5A8AC6"], "Blue - White - Red"],
  [["#FCFCFF", "#63BE7B"], "White - Green"],
  [["#FCFCFF", "#F8696B"], "White - Red"],
  [["#FFEF9C", "#63BE7B"], "Yellow - Green"],
];

/** Data Bars and Color Scales galleries. */
export function CfGalleryDialog(props: { doc: SheetsDocument; sheet: number; sel: Rect; gallery: "dataBar" | "colorScale"; onDone: () => void; onClose: () => void }) {
  const { doc, sheet, sel } = props;
  function add(args: string[]) {
    doc.addCfRule(sheet, sel.top, sel.left, sel.bottom, sel.right, props.gallery, args, "");
    props.onDone();
    props.onClose();
  }
  return (
    <Dialog title={props.gallery === "dataBar" ? "Data Bars" : "Color Scales"} onClose={props.onClose} width={420} footer={button("Cancel", props.onClose)}>
      <div className="cf-gallery">
        {props.gallery === "dataBar"
          ? DATA_BARS.map(([color, name]) => (
              <button key={color} className="cf-swatch" title={`${name} Data Bar`} aria-label={`${name} Data Bar`} onClick={() => add([color])}>
                {[0.9, 0.6, 0.35].map((w) => (
                  <span key={w} className="cf-bar" style={{ width: `${w * 100}%`, background: `linear-gradient(90deg, ${color}, ${color}33)` }} />
                ))}
              </button>
            ))
          : COLOR_SCALES.map(([colors, name]) => (
              <button key={name} className="cf-swatch" title={name} aria-label={name} onClick={() => add(colors)}>
                {colors.map((c, i) => (
                  <span key={i} className="cf-scale" style={{ background: c }} />
                ))}
              </button>
            ))}
      </div>
    </Dialog>
  );
}

interface CfRuleInfo {
  description: string;
  ranges: string;
  format: { fill?: string; color?: string; b?: number; i?: number };
  stopIfTrue: boolean;
  scale?: string[];
  bar?: string;
}

/** Conditional Formatting Rules Manager: reorder and delete this sheet's rules. */
export function CfManageDialog(props: { doc: SheetsDocument; sheet: number; onDone: () => void; onClose: () => void }) {
  const { doc, sheet } = props;
  const rules = useMemo(() => doc.cfRules(sheet).map((j) => JSON.parse(j) as CfRuleInfo), [doc, sheet]);
  // order[i] = index into `rules` of the rule shown at position i.
  const [order, setOrder] = useState(() => rules.map((_, i) => i));
  const [current, setCurrent] = useState(0);
  function move(d: number) {
    const j = current + d;
    if (j < 0 || j >= order.length) return;
    setOrder((o) => {
      const next = [...o];
      [next[current], next[j]] = [next[j], next[current]];
      return next;
    });
    setCurrent(j);
  }
  function ok() {
    const unchanged = order.length === rules.length && order.every((v, i) => v === i);
    if (!unchanged) {
      doc.setCfOrder(sheet, Uint32Array.from(order));
      props.onDone();
    }
    props.onClose();
  }
  const preview = (r: CfRuleInfo) => {
    if (r.bar) return <span className="cf-preview" style={{ background: `linear-gradient(90deg, ${r.bar} 60%, transparent 60%)` }} />;
    if (r.scale) return <span className="cf-preview" style={{ background: `linear-gradient(90deg, ${r.scale.join(",")})` }} />;
    return (
      <span className="cf-preview" style={{ background: r.format.fill ?? "#fff", color: r.format.color ?? "#000", fontWeight: r.format.b ? 700 : 400, fontStyle: r.format.i ? "italic" : "normal" }}>
        AaBbCc
      </span>
    );
  };
  return (
    <Dialog title="Conditional Formatting Rules Manager" onClose={props.onClose} width={600} footer={<>{button("OK", ok, { primary: true })}{button("Cancel", props.onClose)}</>}>
      <div className="dialog-row toolbar">
        <button className="dialog-button" disabled={order.length === 0} onClick={() => {
          setOrder((o) => o.filter((_, i) => i !== current));
          setCurrent((c) => Math.max(0, Math.min(c, order.length - 2)));
        }}>
          Delete Rule
        </button>
        <button className="dialog-button icon" aria-label="Move up" disabled={current === 0} onClick={() => move(-1)}>▲</button>
        <button className="dialog-button icon" aria-label="Move down" disabled={current >= order.length - 1} onClick={() => move(1)}>▼</button>
      </div>
      {order.length === 0 ? (
        <p className="dialog-message">There are no conditional formatting rules on this sheet.</p>
      ) : (
        <table className="dialog-table">
          <thead>
            <tr>
              <th>Rule (applied in order shown)</th>
              <th>Format</th>
              <th>Applies to</th>
            </tr>
          </thead>
          <tbody>
            {order.map((ri, i) => (
              <tr key={ri} className={i === current ? "current" : undefined} onMouseDown={() => setCurrent(i)}>
                <td>{rules[ri].description}</td>
                <td>{preview(rules[ri])}</td>
                <td>={rules[ri].ranges}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </Dialog>
  );
}
