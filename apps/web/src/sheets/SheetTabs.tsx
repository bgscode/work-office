import { useEffect, useRef, useState } from "react";
import { Icon, Tooltip } from "@office/ui";

interface Props {
  names: string[];
  active: number;
  onSelect: (index: number) => void;
  onAdd: () => void;
  /** Returns false if the name was rejected (empty or duplicate). */
  onRename: (index: number, name: string) => boolean;
  onContextMenu: (index: number, x: number, y: number) => void;
  /** Set by the parent to start renaming a tab (e.g. from the tab menu). */
  renaming: number | null;
  setRenaming: (index: number | null) => void;
}

/** Excel-style sheet tabs: scroll arrows, the tabs, and ⊕ to add a sheet. */
export function SheetTabs({ names, active, onSelect, onAdd, onRename, onContextMenu, renaming, setRenaming }: Props) {
  const [draft, setDraft] = useState("");
  const stripRef = useRef<HTMLDivElement>(null);
  const [overflow, setOverflow] = useState({ left: false, right: false });

  function measure() {
    const el = stripRef.current;
    if (!el) return;
    const next = { left: el.scrollLeft > 0, right: el.scrollLeft + el.clientWidth < el.scrollWidth - 1 };
    setOverflow((o) => (o.left === next.left && o.right === next.right ? o : next));
  }

  useEffect(() => {
    const el = stripRef.current!;
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, []);
  useEffect(measure, [names.length]);

  // Keep the active tab visible.
  useEffect(() => {
    stripRef.current?.querySelector<HTMLElement>('[aria-selected="true"]')?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [active]);

  function scrollTabs(direction: 1 | -1) {
    const el = stripRef.current;
    if (!el) return;
    el.scrollBy({ left: direction * Math.max(80, el.clientWidth * 0.6), behavior: "smooth" });
  }

  return (
    <div className="sheet-tabs-bar">
      <div className="sheet-tab-scroll">
        <button aria-label="Scroll tabs left" disabled={!overflow.left} onMouseDown={(e) => e.preventDefault()} onClick={() => scrollTabs(-1)}>
          <Icon name="chevronLeft" size={12} />
        </button>
        <button aria-label="Scroll tabs right" disabled={!overflow.right} onMouseDown={(e) => e.preventDefault()} onClick={() => scrollTabs(1)}>
          <Icon name="chevronRight" size={12} />
        </button>
      </div>
      <div className="sheet-tabs" role="tablist" aria-label="Sheets" ref={stripRef} onScroll={measure}>
        {names.map((name, i) =>
          renaming === i ? (
            <input
              key={i}
              className="tab-rename"
              autoFocus
              value={draft}
              aria-label="Sheet name"
              onFocus={(e) => e.target.select()}
              onChange={(e) => setDraft(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && onRename(i, draft)) setRenaming(null);
                if (e.key === "Escape") setRenaming(null);
              }}
              onBlur={() => {
                onRename(i, draft);
                setRenaming(null);
              }}
            />
          ) : (
            <button
              key={i}
              role="tab"
              aria-selected={i === active}
              className={i === active ? "active" : ""}
              title="Double-click to rename"
              // Keep focus where it is: switching sheets while typing a formula points at that sheet.
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onSelect(i)}
              onDoubleClick={() => {
                setDraft(name);
                setRenaming(i);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                setDraft(name);
                onContextMenu(i, e.clientX, e.clientY);
              }}
            >
              {name}
            </button>
          ),
        )}
      </div>
      <Tooltip label="New sheet" shortcut="Shift+F11">
        <button className="add-sheet" onClick={onAdd} aria-label="Add sheet">
          <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
            <circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor" strokeWidth="1" />
            <path d="M8 5v6M5 8h6" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
          </svg>
        </button>
      </Tooltip>
    </div>
  );
}
