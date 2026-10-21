// Read-only analysis of formula text for the editor: which references it
// contains (for coloured highlights), what the user is typing (for
// autocomplete), and whether a click should insert a reference.

export interface RefMatch {
  start: number;
  end: number;
  /** Zero-based, inclusive. */
  top: number;
  left: number;
  bottom: number;
  right: number;
}

const REF = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseRef(text: string): { row: number; col: number } | null {
  const m = REF.exec(text);
  if (!m) return null;
  const row = Number(m[2]) - 1;
  const col = colIndex(m[1]);
  if (row < 0 || row >= 1_048_576 || col >= 16_384) return null;
  return { row, col };
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

/** Every cell reference or range in a formula (including the leading "="). */
export function findRefs(formula: string): RefMatch[] {
  if (!formula.startsWith("=")) return [];
  const mask = codeMask(formula);
  const out: RefMatch[] = [];
  const word = /[A-Za-z_$][A-Za-z0-9_$.]*(?::[A-Za-z_$][A-Za-z0-9_$.]*)?/g;
  for (let m = word.exec(formula); m; m = word.exec(formula)) {
    const start = m.index;
    const end = start + m[0].length;
    if (!mask[start] || /[A-Za-z0-9_.]/.test(formula[start - 1] ?? "")) continue;
    if (/^\s*\(/.test(formula.slice(end))) continue; // function name
    const [a, b] = m[0].split(":");
    const ra = parseRef(a);
    const rb = b === undefined ? ra : parseRef(b);
    if (!ra || !rb) continue;
    out.push({
      start,
      end,
      top: Math.min(ra.row, rb.row),
      left: Math.min(ra.col, rb.col),
      bottom: Math.max(ra.row, rb.row),
      right: Math.max(ra.col, rb.col),
    });
  }
  return out;
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
