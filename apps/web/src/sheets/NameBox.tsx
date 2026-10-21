import { forwardRef, useEffect, useRef, useState } from "react";
import { Icon } from "@office/ui";
import { findRefs } from "./formulaText";
import type { Selection } from "./geometry";

interface Props {
  /** What to show when not typing, e.g. "B3" or "A1:C5". */
  label: string;
  /** Go to a cell or range. */
  onGo: (s: Selection) => void;
  onDone: () => void;
}

/** Parse a Name Box entry (B12, a1:c5) into a selection, or null. */
export function parseAddress(text: string): Selection | null {
  const t = text.trim();
  const refs = findRefs("=" + t);
  if (refs.length !== 1 || refs[0].start !== 1 || refs[0].end !== t.length + 1) return null;
  const r = refs[0];
  return { anchor: { row: r.top, col: r.left }, focus: { row: r.bottom, col: r.right } };
}

/**
 * The box left of the formula bar. Type a cell or range (B12, A1:C5) and press
 * Enter to go there. Its dropdown lists the places recently gone to.
 */
export const NameBox = forwardRef<HTMLInputElement, Props>(function NameBox({ label, onGo, onDone }, ref) {
  const [draft, setDraft] = useState<string | null>(null);
  const [invalid, setInvalid] = useState(false);
  const [recent, setRecent] = useState<string[]>([]);
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (!wrapRef.current?.contains(e.target as Node)) setOpen(false);
    };
    window.addEventListener("mousedown", onDown, true);
    return () => window.removeEventListener("mousedown", onDown, true);
  }, [open]);

  function goTo(text: string): boolean {
    const s = parseAddress(text);
    if (!s) return false;
    const name = text.trim().toUpperCase();
    setRecent((r) => [name, ...r.filter((x) => x !== name)].slice(0, 8));
    onGo(s);
    return true;
  }

  function go() {
    if (!goTo(draft ?? "")) {
      setInvalid(true);
      return;
    }
    setDraft(null);
    setInvalid(false);
    onDone();
  }

  return (
    <div className="name-box-wrap" ref={wrapRef}>
      <input
        ref={ref}
        className={invalid ? "name-box invalid" : "name-box"}
        aria-label="Name Box: type a cell like B12 and press Enter"
        title="Go to a cell: type B12 or A1:C5 and press Enter (Ctrl+G)"
        spellCheck={false}
        value={draft ?? label}
        onFocus={(e) => {
          // Select the current address so typing replaces it.
          setDraft(label);
          e.target.select();
        }}
        onChange={(e) => {
          setDraft(e.target.value);
          setInvalid(false);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            setOpen(false);
            go();
          } else if (e.key === "Escape") {
            e.preventDefault();
            setOpen(false);
            setDraft(null);
            setInvalid(false);
            onDone();
          } else if (e.key === "ArrowDown" && e.altKey) {
            e.preventDefault();
            setOpen(true);
          }
        }}
        onBlur={() => {
          setDraft(null);
          setInvalid(false);
        }}
      />
      <button
        className="name-box-arrow"
        aria-label="Recent places"
        aria-haspopup="listbox"
        aria-expanded={open}
        tabIndex={-1}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((o) => !o)}
      >
        <Icon name="chevronDown" size={10} />
      </button>
      {open && (
        <ul className="name-box-list office-menu" role="listbox" aria-label="Recent places">
          {recent.length === 0 && <li className="empty">Type a cell (B12) or range (A1:C5) in the box and press Enter.</li>}
          {recent.map((r) => (
            <li
              key={r}
              role="option"
              aria-selected={false}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                setOpen(false);
                goTo(r);
                onDone();
              }}
            >
              {r}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
});
