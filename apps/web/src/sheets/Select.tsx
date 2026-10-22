// Excel-style dropdown used everywhere instead of the browser's <select>, so
// ribbon boxes, dialogs and the status bar look and behave the same on every
// platform. It takes <option> children and an onChange like <select>'s, so it
// drops in where a native select was. The list opens in a portal positioned
// under the box (above it when there is no room), so it is never clipped by
// the ribbon or a dialog.

import {
  Children,
  isValidElement,
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type ReactElement,
  type ReactNode,
  type RefObject,
} from "react";
import { createPortal } from "react-dom";

interface Option {
  value: string;
  label: ReactNode;
  disabled: boolean;
}

export interface SelectProps {
  value: string | number;
  /** Same shape as a native select's change event, so handlers carry over. */
  onChange: (e: { target: { value: string } }) => void;
  children: ReactNode;
  "aria-label"?: string;
  title?: string;
  className?: string;
  style?: CSSProperties;
  autoFocus?: boolean;
  disabled?: boolean;
}

function readOptions(children: ReactNode): Option[] {
  const out: Option[] = [];
  Children.forEach(children, (child) => {
    if (!isValidElement(child)) return;
    const el = child as ReactElement<{ value?: string | number; children?: ReactNode; disabled?: boolean }>;
    if (el.type !== "option") return;
    out.push({ value: String(el.props.value ?? ""), label: el.props.children, disabled: !!el.props.disabled });
  });
  return out;
}

/** Fixed-position style for a popup list under (or above) `anchor`. */
export function usePopupPosition(anchor: RefObject<HTMLElement>, open: boolean, maxHeight = 300): CSSProperties | null {
  const [style, setStyle] = useState<CSSProperties | null>(null);
  useLayoutEffect(() => {
    if (!open) {
      setStyle(null);
      return;
    }
    const place = () => {
      const r = anchor.current?.getBoundingClientRect();
      if (!r) return;
      const below = window.innerHeight - r.bottom - 8;
      const above = r.top - 8;
      const up = below < Math.min(maxHeight, 160) && above > below;
      const height = Math.min(maxHeight, up ? above : below);
      setStyle({
        position: "fixed",
        left: Math.min(r.left, window.innerWidth - r.width - 4),
        minWidth: r.width,
        maxHeight: height,
        ...(up ? { bottom: window.innerHeight - r.top + 1 } : { top: r.bottom + 1 }),
      });
    };
    place();
    window.addEventListener("resize", place);
    window.addEventListener("scroll", place, true);
    return () => {
      window.removeEventListener("resize", place);
      window.removeEventListener("scroll", place, true);
    };
  }, [open, anchor, maxHeight]);
  return style;
}

export function Select({ value, onChange, children, title, className, style, autoFocus, disabled, ...rest }: SelectProps) {
  const options = readOptions(children);
  const current = options.findIndex((o) => o.value === String(value));
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(current);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const listRef = useRef<HTMLUListElement>(null);
  const listId = useId();
  const popup = usePopupPosition(buttonRef, open);

  useEffect(() => {
    if (autoFocus) buttonRef.current?.focus();
  }, []);

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      const t = e.target as Node;
      if (!buttonRef.current?.contains(t) && !listRef.current?.contains(t)) setOpen(false);
    };
    window.addEventListener("mousedown", onDown, true);
    return () => window.removeEventListener("mousedown", onDown, true);
  }, [open]);

  useEffect(() => {
    if (open && active >= 0) listRef.current?.children[active]?.scrollIntoView({ block: "nearest" });
  }, [open, active, popup]);

  function show() {
    setActive(Math.max(0, current));
    setOpen(true);
  }

  function pick(i: number) {
    const o = options[i];
    setOpen(false);
    buttonRef.current?.focus();
    if (o && !o.disabled && o.value !== String(value)) onChange({ target: { value: o.value } });
  }

  function step(from: number, dir: 1 | -1): number {
    for (let i = from + dir; i >= 0 && i < options.length; i += dir) if (!options[i].disabled) return i;
    return from;
  }

  const label = current >= 0 ? options[current].label : "";
  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        className={className ? `xl-select ${className}` : "xl-select"}
        style={style}
        title={title}
        disabled={disabled}
        role="combobox"
        aria-label={rest["aria-label"]}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-controls={open ? listId : undefined}
        onMouseDown={(e) => {
          e.preventDefault();
          buttonRef.current?.focus();
          if (open) setOpen(false);
          else show();
        }}
        onKeyDown={(e) => {
          if (!open) {
            if (e.key === "ArrowDown" || e.key === "ArrowUp" || e.key === "Enter" || e.key === " ") {
              e.preventDefault();
              show();
            }
            return;
          }
          if (e.key === "ArrowDown" || e.key === "ArrowUp") setActive((a) => step(a, e.key === "ArrowDown" ? 1 : -1));
          else if (e.key === "Home") setActive(step(-1, 1));
          else if (e.key === "End") setActive(step(options.length, -1));
          else if (e.key === "Enter" || e.key === " ") pick(active);
          else if (e.key === "Escape" || e.key === "Tab") {
            if (e.key === "Escape") e.stopPropagation();
            setOpen(false);
            if (e.key === "Tab") return;
          } else return;
          e.preventDefault();
        }}
      >
        <span className="xl-select-text">{label}</span>
        <svg className="xl-select-arrow" width="8" height="5" viewBox="0 0 8 5" aria-hidden>
          <path d="M0 0h8L4 5z" fill="currentColor" />
        </svg>
      </button>
      {open &&
        popup &&
        createPortal(
          <ul ref={listRef} id={listId} role="listbox" aria-label={rest["aria-label"]} className="xl-popup-list" style={popup}>
            {options.map((o, i) => (
              <li
                key={o.value}
                role="option"
                aria-selected={i === current}
                aria-disabled={o.disabled || undefined}
                className={[i === active ? "active" : "", i === current ? "current" : "", o.disabled ? "disabled" : ""].join(" ").trim()}
                onMouseEnter={() => !o.disabled && setActive(i)}
                onMouseDown={(e) => {
                  e.preventDefault();
                  pick(i);
                }}
              >
                {o.label}
              </li>
            ))}
          </ul>,
          document.body,
        )}
    </>
  );
}
