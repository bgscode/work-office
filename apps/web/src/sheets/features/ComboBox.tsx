// An editable dropdown like Excel's Font and Font Size boxes: type a value and
// press Enter, or open the list (click the arrow, Alt+Down or Down) and pick.
// Escape or clicking away restores the current value.

import { useEffect, useId, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { usePopupPosition } from "../Select";

export interface ComboBoxProps {
  label: string;
  /** Value shown when not being edited (the active cell's). */
  value: string;
  options: string[];
  /** Called with the typed or picked text. */
  onCommit: (value: string) => void;
  /** Return keyboard focus to the grid. */
  onDone: () => void;
  /** Width of the text box in pixels. */
  width: number;
  /** Custom rendering of a list entry (e.g. the font name in its own font). */
  renderOption?: (option: string) => ReactNode;
  title?: string;
}

/** The option to highlight for typed text: an exact match, else the first that starts with it. */
export function bestMatch(options: string[], typed: string): number {
  const t = typed.trim().toLowerCase();
  if (!t) return -1;
  const exact = options.findIndex((o) => o.toLowerCase() === t);
  return exact >= 0 ? exact : options.findIndex((o) => o.toLowerCase().startsWith(t));
}

export function ComboBox({ label, value, options, onCommit, onDone, width, renderOption, title }: ComboBoxProps) {
  const [text, setText] = useState<string | null>(null);
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(-1);
  const inputRef = useRef<HTMLInputElement>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const listRef = useRef<HTMLUListElement>(null);
  const listId = useId();
  const shown = text ?? value;
  // The list opens in a portal so the ribbon's overflow never clips it.
  const popup = usePopupPosition(rootRef, open, 320);

  // Close when clicking elsewhere.
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      const t = e.target as Node;
      if (!rootRef.current?.contains(t) && !listRef.current?.contains(t)) close(false);
    };
    window.addEventListener("mousedown", onDown, true);
    return () => window.removeEventListener("mousedown", onDown, true);
  }, [open]);

  useEffect(() => {
    if (open && active >= 0) listRef.current?.children[active]?.scrollIntoView({ block: "nearest" });
  }, [open, active, popup]);

  function openList() {
    setOpen(true);
    setActive(Math.max(0, bestMatch(options, shown)));
  }

  function close(refocus: boolean) {
    setOpen(false);
    setText(null);
    setActive(-1);
    if (refocus) onDone();
  }

  function commit(v: string) {
    setOpen(false);
    setText(null);
    setActive(-1);
    if (v.trim() && v.trim() !== value) onCommit(v.trim());
    onDone();
  }

  return (
    <div className="fmt-combo" ref={rootRef} title={title ?? label}>
      <input
        ref={inputRef}
        role="combobox"
        aria-label={label}
        aria-expanded={open}
        aria-controls={listId}
        aria-autocomplete="list"
        aria-activedescendant={open && active >= 0 ? `${listId}-${active}` : undefined}
        style={{ width }}
        value={shown}
        spellCheck={false}
        onFocus={(e) => e.currentTarget.select()}
        onChange={(e) => {
          setText(e.target.value);
          const m = bestMatch(options, e.target.value);
          setActive(m);
          if (m >= 0) setOpen(true);
        }}
        onBlur={(e) => {
          const to = e.relatedTarget as Node | null;
          if (!rootRef.current?.contains(to) && !listRef.current?.contains(to)) {
            setOpen(false);
            setText(null);
          }
        }}
        onKeyDown={(e) => {
          if (e.key === "ArrowDown" || e.key === "ArrowUp") {
            e.preventDefault();
            if (!open || e.altKey) {
              openList();
              return;
            }
            const next = e.key === "ArrowDown" ? Math.min(options.length - 1, active + 1) : Math.max(0, active - 1);
            setActive(next);
            setText(options[next]);
          } else if (e.key === "Enter" || e.key === "Tab") {
            if (e.key === "Enter") e.preventDefault();
            const typed = text ?? value;
            // A typed prefix of a listed font picks that font, like Excel's autocomplete.
            commit(open && active >= 0 && text !== null && options[active].toLowerCase().startsWith(typed.trim().toLowerCase()) ? options[active] : typed);
          } else if (e.key === "Escape") {
            e.preventDefault();
            close(true);
          }
        }}
      />
      <button
        type="button"
        className="fmt-combo-arrow"
        tabIndex={-1}
        aria-label={`Show ${label} list`}
        onMouseDown={(e) => {
          e.preventDefault();
          if (open) close(false);
          else {
            inputRef.current?.focus();
            openList();
          }
        }}
      >
        <svg width="8" height="5" viewBox="0 0 8 5" aria-hidden>
          <path d="M0 0h8L4 5z" fill="currentColor" />
        </svg>
      </button>
      {open &&
        popup &&
        createPortal(
        <ul ref={listRef} id={listId} role="listbox" aria-label={label} className="xl-popup-list" style={popup}>
          {options.map((o, i) => (
            <li
              key={o}
              id={`${listId}-${i}`}
              role="option"
              aria-selected={i === active}
              className={i === active ? "active" : o === value ? "current" : ""}
              onMouseDown={(e) => {
                e.preventDefault();
                commit(o);
              }}
              onMouseEnter={() => setActive(i)}
            >
              {renderOption ? renderOption(o) : o}
            </li>
          ))}
        </ul>,
          document.body,
        )}
    </div>
  );
}
