// Read-only analysis of formula text for the editor: which references it
// contains (for coloured highlights), what the user is typing (for
// autocomplete), and whether a click should insert a reference.

export interface RefMatch {
  start: number;
  end: number;
  /** Sheet named before `!` (unquoted), or undefined for the formula's own sheet. */
  sheet?: string;
  /** Zero-based, inclusive. Whole columns/rows span every row/column. */
  top: number;
  left: number;
  bottom: number;
  right: number;
}

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Positions outside string literals. */
function codeMask(formula: string): boolean[] {
  const mask: boolean[] = [];
  let inString = false;
  for (const ch of formula) {
    if (ch === '"') {
      mask.push(false);
      inString = !inString;
    } else {
      mask.push(!inString);
    }
  }
  return mask;
}

// Pieces of a reference, matched at a fixed position (sticky regexes).
const QUOTED_SHEET = /'((?:[^']|'')+)'!/y;
const PLAIN_SHEET = /([A-Za-z_\u0080-\uffff][A-Za-z0-9_.\u0080-\uffff]*)!/y;
const CELL_RANGE = /\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?/y;
const COL_RANGE = /\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})(?![A-Za-z0-9])/y;
const ROW_RANGE = /\$?(\d+):\$?(\d+)(?![0-9.])/y;
const IDENT = /[A-Za-z0-9_.$]/;

function stickyAt(re: RegExp, text: string, i: number): RegExpExecArray | null {
  re.lastIndex = i;
  return re.exec(text);
}

/**
 * Every reference in a formula (including the leading "="): cells and ranges
 * (`A1`, `$B$2:C9`), whole columns and rows (`A:A`, `3:5`), with an optional
 * sheet (`Sheet2!A1`, `'My Sheet'!B:B`). Skips strings and function names.
 */
export function findRefs(formula: string): RefMatch[] {
  if (!formula.startsWith("=")) return [];
  const mask = codeMask(formula);
  const out: RefMatch[] = [];
  let i = 1;
  while (i < formula.length) {
    if (!mask[i] || IDENT.test(formula[i - 1] ?? "") || formula[i - 1] === "'") {
      i++;
      continue;
    }
    const start = i;
    let sheet: string | undefined;
    let at = i;
    const quoted = stickyAt(QUOTED_SHEET, formula, at);
    const plain = quoted ? null : stickyAt(PLAIN_SHEET, formula, at);
    if (quoted) {
      sheet = quoted[1].replace(/''/g, "'");
      at += quoted[0].length;
    } else if (plain) {
      sheet = plain[1];
      at += plain[0].length;
    }
    const cell = stickyAt(CELL_RANGE, formula, at);
    const cols = cell ? null : stickyAt(COL_RANGE, formula, at);
    const rows = cell || cols ? null : stickyAt(ROW_RANGE, formula, at);
    const m = cell ?? cols ?? rows;
    const end = m ? at + m[0].length : at;
    // Not a reference, or the name of a function (followed by "(") or a longer word.
    if (!m || IDENT.test(formula[end] ?? "") || /^\s*\(/.test(formula.slice(end))) {
      i = Math.max(i + 1, sheet ? at : i + 1);
      continue;
    }
    let top: number, left: number, bottom: number, right: number;
    if (cell) {
      const r1 = Number(cell[2]) - 1;
      const c1 = colIndex(cell[1]);
      const r2 = cell[4] ? Number(cell[4]) - 1 : r1;
      const c2 = cell[3] ? colIndex(cell[3]) : c1;
      [top, bottom, left, right] = [Math.min(r1, r2), Math.max(r1, r2), Math.min(c1, c2), Math.max(c1, c2)];
    } else if (cols) {
      const c1 = colIndex(cols[1]);
      const c2 = colIndex(cols[2]);
      [top, bottom, left, right] = [0, MAX_ROW, Math.min(c1, c2), Math.max(c1, c2)];
    } else {
      const r1 = Number(rows![1]) - 1;
      const r2 = Number(rows![2]) - 1;
      [top, bottom, left, right] = [Math.min(r1, r2), Math.max(r1, r2), 0, MAX_COL];
    }
    if (top < 0 || bottom > MAX_ROW || left < 0 || right > MAX_COL) {
      i = end;
      continue;
    }
    out.push({ start, end, sheet, top, left, bottom, right });
    i = end;
  }
  return out;
}

/** References that point at the sheet being viewed (for coloured highlights). */
export function refsOnSheet(refs: RefMatch[], viewed: string, editedOn: string): RefMatch[] {
  const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();
  return refs.filter((r) => (r.sheet === undefined ? same(viewed, editedOn) : same(r.sheet, viewed)));
}

/**
 * F4: cycle the reference at (or just before) the caret through
 * `$A$1` → `A$1` → `$A1` → `A1` → `$A$1`. Returns the new formula and caret,
 * or null when the caret isn't on a reference.
 */
export function cycleReference(formula: string, caret: number): { formula: string; caret: number } | null {
  const ref = findRefs(formula).find((r) => caret >= r.start && caret <= r.end);
  if (!ref) return null;
  const text = formula.slice(ref.start, ref.end);
  const bang = text.lastIndexOf("!");
  const prefix = bang >= 0 ? text.slice(0, bang + 1) : "";
  const body = bang >= 0 ? text.slice(bang + 1) : text;
  const parts = body.split(":").map((p) => {
    const m = /^(\$?)([A-Za-z]*)(\$?)(\d*)$/.exec(p);
    return m ? { col: m[2], row: m[4], absCol: m[1] === "$", absRow: m[3] === "$" } : null;
  });
  if (parts.some((p) => !p)) return null;
  const first = parts[0]!;
  // Next state from the first part: both → row only → column only → none → both.
  const [absCol, absRow] =
    first.absCol && first.absRow ? [false, true] : !first.absCol && first.absRow ? [true, false] : first.absCol ? [false, false] : [true, true];
  const write = (p: { col: string; row: string }) =>
    `${p.col ? (absCol ? "$" : "") + p.col : ""}${p.row ? (absRow ? "$" : "") + p.row : ""}`;
  const next = prefix + parts.map((p) => write(p!)).join(":");
  return { formula: formula.slice(0, ref.start) + next + formula.slice(ref.end), caret: ref.start + next.length };
}

/** The function-name prefix being typed at `caret`, e.g. "VLO" in "=SUM(1)+VLO". */
export function wordAtCaret(formula: string, caret: number): { start: number; text: string } | null {
  if (!formula.startsWith("=")) return null;
  const before = formula.slice(0, caret);
  if (!codeMask(formula + " ")[caret]) return null; // inside a string
  const m = /[A-Za-z_][A-Za-z0-9_.]*$/.exec(before);
  if (!m) return null;
  if (/[$0-9]/.test(before[m.index - 1] ?? "")) return null;
  return { start: m.index, text: m[0] };
}

/** Name of the innermost function call whose parentheses contain `caret`. */
export function enclosingFunction(formula: string, caret: number): string | null {
  if (!formula.startsWith("=")) return null;
  const mask = codeMask(formula);
  let depth = 0;
  for (let i = caret - 1; i > 0; i--) {
    if (!mask[i]) continue;
    const ch = formula[i];
    if (ch === ")") depth++;
    else if (ch === "(") {
      if (depth === 0) {
        const m = /([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(formula.slice(0, i));
        return m ? m[1].toUpperCase() : null;
      }
      depth--;
    }
  }
  return null;
}

/** True when a reference may be inserted at `caret` (right after "=", "(", ",", or an operator). */
export function canInsertRef(formula: string, caret: number): boolean {
  if (!formula.startsWith("=") || caret === 0) return false;
  const before = formula.slice(0, caret).trimEnd();
  if (!codeMask(formula + " ")[caret]) return false;
  return /[=(,+\-*/^&<>:%]$/.test(before);
}

/** Close any parentheses the user left open, as Excel does on Enter. */
export function closeParens(formula: string): string {
  if (!formula.startsWith("=")) return formula;
  const mask = codeMask(formula);
  let open = 0;
  for (let i = 0; i < formula.length; i++) {
    if (!mask[i]) continue;
    if (formula[i] === "(") open++;
    else if (formula[i] === ")") open = Math.max(0, open - 1);
  }
  return formula + ")".repeat(open);
}
