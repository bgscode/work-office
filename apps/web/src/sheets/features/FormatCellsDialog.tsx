// Format Cells (Ctrl+1), like Excel's: Number (categories, decimals,
// separators, negative numbers, currency symbols, dates, times, fractions,
// custom codes with a live sample), Alignment, Font, Border and Fill.
// OK applies every change as one undo step (`formatRangeMany`).

import { useMemo, useState, type ReactNode } from "react";
import { SheetsDocument, type CellStyle } from "@office/engine";
import { Dialog } from "../Dialog";
import { Select } from "../Select";
import { DEFAULT_FONT, DEFAULT_FONT_SIZE, FONT_CHOICES, FONT_SIZE_CHOICES } from "../cellRenderer";
import {
  CATEGORIES,
  CUSTOM_CODES,
  DATE_CODES,
  FRACTION_CODES,
  NEGATIVE_STYLES,
  SYMBOLS,
  TIME_CODES,
  accountingCode,
  categoryOf,
  currencyCode,
  decimalsOf,
  numberCode,
  percentCode,
  scientificCode,
  type Category,
  type NegativeStyle,
} from "./numberFormats";

export type FormatTab = "number" | "alignment" | "font" | "border" | "fill";

interface Props {
  style: CellStyle;
  /** What the active cell holds, for the sample ("" when empty). */
  sample: string;
  initialTab?: FormatTab;
  /** `formatRange` keys and values, applied together. */
  onApply: (changes: [string, string][]) => void;
  onBorders: (preset: string, style: string, color: string) => void;
  onClose: () => void;
}

export const FILL_COLORS = [
  "#FFFFFF", "#F2F2F2", "#D9D9D9", "#BFBFBF", "#808080", "#000000",
  "#FFF2CC", "#FFE699", "#FFC000", "#F8CBAD", "#F4B183", "#ED7D31",
  "#E2EFDA", "#C6E0B4", "#70AD47", "#DDEBF7", "#BDD7EE", "#4472C4",
  "#FCE4D6", "#FF0000", "#C00000", "#E4DFEC", "#7030A0", "#002060",
];

export function FormatCellsDialog({ style, sample, initialTab = "number", onApply, onBorders, onClose }: Props) {
  const startCode = (style as CellStyle & { code?: string }).code ?? "General";
  const [tab, setTab] = useState<FormatTab>(initialTab);
  // Number tab
  const [category, setCategory] = useState<Category>(categoryOf(startCode, style.fmt));
  const [decimals, setDecimals] = useState(category === "general" ? 2 : decimalsOf(startCode));
  const [thousands, setThousands] = useState(startCode.includes(","));
  const [negative, setNegative] = useState<NegativeStyle>(startCode.includes("[Red]") ? (startCode.includes("(") ? "redParens" : "red") : startCode.includes("(") ? "parens" : "minus");
  const [symbol, setSymbol] = useState(startCode.includes("₹") || !/[$€£¥]/.test(startCode) ? "inr" : startCode.includes("€") ? "eur" : startCode.includes("£") ? "gbp" : startCode.includes("¥") ? "jpy" : "usd");
  const [dateCode, setDateCode] = useState(DATE_CODES.includes(startCode) ? startCode : DATE_CODES[0]);
  const [timeCode, setTimeCode] = useState(TIME_CODES.includes(startCode) ? startCode : TIME_CODES[0]);
  const [fractionCode, setFractionCode] = useState(FRACTION_CODES.some((f) => f.code === startCode) ? startCode : FRACTION_CODES[0].code);
  const [customCode, setCustomCode] = useState(startCode);
  // Other tabs
  const [align, setAlign] = useState<string>(style.align ?? "general");
  const [valign, setValign] = useState<string>(style.va ?? "bottom");
  const [wrap, setWrap] = useState(!!style.wrap);
  const [font, setFont] = useState(style.font ?? DEFAULT_FONT);
  const [size, setSize] = useState(String(style.sz ?? DEFAULT_FONT_SIZE));
  const [bold, setBold] = useState(!!style.b);
  const [italic, setItalic] = useState(!!style.i);
  const [underline, setUnderline] = useState(!!style.u);
  const [strike, setStrike] = useState(!!style.s);
  const [color, setColor] = useState(style.color ?? "#000000");
  const [fill, setFill] = useState<string | null>(style.fill ?? null);
  const [border, setBorder] = useState<{ preset: string; style: string; color: string } | null>(null);

  const code = useMemo(() => {
    switch (category) {
      case "general":
        return "General";
      case "number":
        return numberCode(decimals, thousands, negative);
      case "currency":
        return currencyCode(decimals, symbol, negative);
      case "accounting":
        return accountingCode(decimals, symbol);
      case "date":
        return dateCode;
      case "time":
        return timeCode;
      case "percentage":
        return percentCode(decimals);
      case "fraction":
        return fractionCode;
      case "scientific":
        return scientificCode(decimals);
      case "text":
        return "@";
      case "custom":
        return customCode;
    }
  }, [category, decimals, thousands, negative, symbol, dateCode, timeCode, fractionCode, customCode]);

  const sampleInput = sample || "1234.5";
  const [previewText, previewColor] = SheetsDocument.previewFormat(sampleInput, code);

  function ok() {
    const changes: [string, string][] = [];
    if (code !== startCode) changes.push(["numfmt", code]);
    if (align !== (style.align ?? "general")) changes.push(["align", align === "general" ? "" : align]);
    if (valign !== (style.va ?? "bottom")) changes.push(["valign", valign]);
    if (wrap !== !!style.wrap) changes.push(["wrap", wrap ? "1" : "0"]);
    if (font !== (style.font ?? DEFAULT_FONT)) changes.push(["font", font === DEFAULT_FONT ? "" : font]);
    if (Number(size) !== (style.sz ?? DEFAULT_FONT_SIZE) && Number(size) > 0) changes.push(["size", size]);
    if (bold !== !!style.b) changes.push(["bold", bold ? "1" : "0"]);
    if (italic !== !!style.i) changes.push(["italic", italic ? "1" : "0"]);
    if (underline !== !!style.u) changes.push(["underline", underline ? "1" : "0"]);
    if (strike !== !!style.s) changes.push(["strike", strike ? "1" : "0"]);
    if (color !== (style.color ?? "#000000")) changes.push(["color", color]);
    if (fill !== (style.fill ?? null)) changes.push(["fill", fill ?? ""]);
    if (changes.length) onApply(changes);
    if (border) onBorders(border.preset, border.style, border.color);
    onClose();
  }

  const check = (label: string, value: boolean, set: (v: boolean) => void) => (
    <label className="dialog-check">
      <input type="checkbox" checked={value} onChange={(e) => set(e.target.checked)} /> {label}
    </label>
  );
  const decimalsField = (
    <label className="dialog-field inline">
      <span>Decimal places:</span>
      <input aria-label="Decimal places" type="number" min={0} max={30} value={decimals} onChange={(e) => setDecimals(Math.max(0, Math.min(30, Number(e.target.value) || 0)))} style={{ width: 64 }} />
    </label>
  );
  const symbolField = (
    <label className="dialog-field inline">
      <span>Symbol:</span>
      <Select aria-label="Currency symbol" value={symbol} onChange={(e) => setSymbol(e.target.value)}>
        {SYMBOLS.map((s) => (
          <option key={s.id} value={s.id}>
            {s.label}
          </option>
        ))}
      </Select>
    </label>
  );
  const listBox = (items: { value: string; label: ReactNode }[], value: string, set: (v: string) => void, label: string) => (
    <ul className="format-list" role="listbox" aria-label={label}>
      {items.map((it) => (
        <li key={it.value} role="option" aria-selected={it.value === value} className={it.value === value ? "active" : undefined} onMouseDown={() => set(it.value)}>
          {it.label}
        </li>
      ))}
    </ul>
  );
  const sampleOf = (c: string) => SheetsDocument.previewFormat(sampleInput, c)[0];

  let numberBody: ReactNode = null;
  if (category === "number" || category === "currency") {
    numberBody = (
      <>
        {decimalsField}
        {category === "number" ? check("Use 1000 Separator (,)", thousands, setThousands) : symbolField}
        <p className="dialog-label">Negative numbers:</p>
        {listBox(
          NEGATIVE_STYLES.map((n) => ({ value: n.id, label: <span style={{ color: n.color }}>{n.label}</span> })),
          negative,
          (v) => setNegative(v as NegativeStyle),
          "Negative numbers",
        )}
      </>
    );
  } else if (category === "accounting") {
    numberBody = (
      <>
        {decimalsField}
        {symbolField}
      </>
    );
  } else if (category === "percentage" || category === "scientific") {
    numberBody = decimalsField;
  } else if (category === "date") {
    numberBody = listBox(DATE_CODES.map((c) => ({ value: c, label: sampleOf(c) || c })), dateCode, setDateCode, "Date type");
  } else if (category === "time") {
    numberBody = listBox(TIME_CODES.map((c) => ({ value: c, label: sampleOf(c) || c })), timeCode, setTimeCode, "Time type");
  } else if (category === "fraction") {
    numberBody = listBox(FRACTION_CODES.map((f) => ({ value: f.code, label: f.label })), fractionCode, setFractionCode, "Fraction type");
  } else if (category === "custom") {
    numberBody = (
      <>
        <label className="dialog-field">
          <span>Type:</span>
          <input aria-label="Format code" value={customCode} onChange={(e) => setCustomCode(e.target.value)} spellCheck={false} />
        </label>
        {listBox(CUSTOM_CODES.map((c) => ({ value: c, label: c })), customCode, setCustomCode, "Format code list")}
      </>
    );
  }

  let body: ReactNode;
  if (tab === "number") {
    body = (
      <div className="format-number">
        {listBox(CATEGORIES.map((c) => ({ value: c.id, label: c.label })), category, (v) => setCategory(v as Category), "Category")}
        <div className="format-number-options">
          <fieldset className="format-sample">
            <legend>Sample</legend>
            <span style={{ color: previewColor || undefined }}>{previewText || " "}</span>
          </fieldset>
          {numberBody}
          <p className="dialog-hint">{CATEGORIES.find((c) => c.id === category)?.description}</p>
        </div>
      </div>
    );
  } else if (tab === "alignment") {
    body = (
      <>
        <p className="dialog-label">Text alignment</p>
        <label className="dialog-field inline">
          <span>Horizontal:</span>
          <Select aria-label="Horizontal alignment" value={align} onChange={(e) => setAlign(e.target.value)}>
            <option value="general">General</option>
            <option value="left">Left</option>
            <option value="center">Center</option>
            <option value="right">Right</option>
          </Select>
        </label>
        <label className="dialog-field inline">
          <span>Vertical:</span>
          <Select aria-label="Vertical alignment" value={valign} onChange={(e) => setValign(e.target.value)}>
            <option value="top">Top</option>
            <option value="center">Center</option>
            <option value="bottom">Bottom</option>
          </Select>
        </label>
        <p className="dialog-label">Text control</p>
        {check("Wrap text", wrap, setWrap)}
      </>
    );
  } else if (tab === "font") {
    body = (
      <>
        <div className="dialog-row">
          <label className="dialog-field" style={{ flex: 2 }}>
            <span>Font:</span>
            <Select aria-label="Font" value={font} onChange={(e) => setFont(e.target.value)}>
              {[...new Set([font, ...FONT_CHOICES])].map((f) => (
                <option key={f} value={f}>
                  <span style={{ fontFamily: f }}>{f}</span>
                </option>
              ))}
            </Select>
          </label>
          <label className="dialog-field" style={{ flex: 1 }}>
            <span>Size:</span>
            <Select aria-label="Size" value={size} onChange={(e) => setSize(e.target.value)}>
              {[...new Set([size, ...FONT_SIZE_CHOICES.map(String)])].map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </Select>
          </label>
        </div>
        <p className="dialog-label">Font style and effects</p>
        <div className="dialog-row">
          {check("Bold", bold, setBold)}
          {check("Italic", italic, setItalic)}
          {check("Underline", underline, setUnderline)}
          {check("Strikethrough", strike, setStrike)}
        </div>
        <label className="dialog-field inline">
          <span>Color:</span>
          <input type="color" aria-label="Font color" value={color} onChange={(e) => setColor(e.target.value)} />
        </label>
        <fieldset className="format-sample">
          <legend>Preview</legend>
          <span
            style={{
              fontFamily: font,
              fontSize: `${Number(size) || 11}pt`,
              fontWeight: bold ? 700 : 400,
              fontStyle: italic ? "italic" : "normal",
              textDecoration: [underline ? "underline" : "", strike ? "line-through" : ""].join(" ").trim() || "none",
              color,
            }}
          >
            AaBbCcYyZz
          </span>
        </fieldset>
      </>
    );
  } else if (tab === "border") {
    const pick = (preset: string) => setBorder({ preset, style: border?.style ?? "thin", color: border?.color ?? "" });
    body = (
      <>
        <p className="dialog-label">Presets</p>
        <div className="dialog-row toolbar">
          {[
            ["none", "None"],
            ["outside", "Outline"],
            ["inside", "Inside"],
            ["all", "All"],
          ].map(([id, label]) => (
            <button key={id} className={border?.preset === id ? "dialog-button primary" : "dialog-button"} onClick={() => pick(id)}>
              {label}
            </button>
          ))}
        </div>
        <label className="dialog-field inline">
          <span>Line style:</span>
          <Select aria-label="Line style" value={border?.style ?? "thin"} onChange={(e) => setBorder({ preset: border?.preset ?? "outside", style: e.target.value, color: border?.color ?? "" })}>
            <option value="thin">Thin</option>
            <option value="medium">Medium</option>
            <option value="thick">Thick</option>
            <option value="dashed">Dashed</option>
            <option value="dotted">Dotted</option>
            <option value="double">Double</option>
          </Select>
        </label>
        <label className="dialog-field inline">
          <span>Color:</span>
          <input type="color" aria-label="Border color" value={border?.color || "#000000"} onChange={(e) => setBorder({ preset: border?.preset ?? "outside", style: border?.style ?? "thin", color: e.target.value })} />
        </label>
        <p className="dialog-hint">Pick a preset, then a line style and colour. The border is applied to the selected cells when you click OK.</p>
      </>
    );
  } else {
    body = (
      <>
        <p className="dialog-label">Background Color</p>
        <button className={fill === null ? "dialog-button primary" : "dialog-button"} onClick={() => setFill(null)}>
          No Color
        </button>
        <div className="fill-swatches" role="listbox" aria-label="Fill colours">
          {FILL_COLORS.map((c) => (
            <button key={c} role="option" aria-selected={fill === c} aria-label={c} className={fill === c ? "active" : undefined} style={{ background: c }} onClick={() => setFill(c)} />
          ))}
        </div>
        <label className="dialog-field inline">
          <span>More colors:</span>
          <input type="color" aria-label="Fill color" value={fill ?? "#FFFFFF"} onChange={(e) => setFill(e.target.value)} />
        </label>
      </>
    );
  }

  const tabs: [FormatTab, string][] = [
    ["number", "Number"],
    ["alignment", "Alignment"],
    ["font", "Font"],
    ["border", "Border"],
    ["fill", "Fill"],
  ];
  return (
    <Dialog
      title="Format Cells"
      onClose={onClose}
      width={560}
      className="format-cells"
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
