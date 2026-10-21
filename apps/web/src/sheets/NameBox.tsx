import { forwardRef, useState } from "react";
import { findRefs } from "./formulaText";
import type { Selection } from "./geometry";

interface Props {
  /** What to show when not typing, e.g. "B3" or "A1:C5". */
  label: string;
  /** Go to a cell or range. */
  onGo: (s: Selection) => void;
  onDone: () => void;
}

/** The box left of the formula bar. Type a cell or range (B12, A1:C5) and press Enter to go there. */
export const NameBox = forwardRef<HTMLInputElement, Props>(function NameBox({ label, onGo, onDone }, ref) {
  const [draft, setDraft] = useState<string | null>(null);
  const [invalid, setInvalid] = useState(false);

  function go() {
    const text = (draft ?? "").trim();
    const refs = findRefs("=" + text);
    if (refs.length !== 1 || refs[0].start !== 1 || refs[0].end !== text.length + 1) {
      setInvalid(true);
      return;
    }
    const r = refs[0];
    onGo({ anchor: { row: r.top, col: r.left }, focus: { row: r.bottom, col: r.right } });
    setDraft(null);
    setInvalid(false);
    onDone();
  }

  return (
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
          go();
        } else if (e.key === "Escape") {
          e.preventDefault();
          setDraft(null);
          setInvalid(false);
          onDone();
        }
      }}
      onBlur={() => {
        setDraft(null);
        setInvalid(false);
      }}
    />
  );
});
