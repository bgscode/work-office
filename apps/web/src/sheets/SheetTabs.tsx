import { useEffect, useRef, useState } from "react";
import { Icon, Tooltip } from "@office/ui";
import { moveTarget, textOn, type SheetTab } from "./features/sheetTabs";

interface Props {
  names: string[];
  /** Colour and hidden state per sheet; hidden sheets get no tab. */
  tabs: SheetTab[];
  active: number;
  onSelect: (index: number) => void;
  onAdd: () => void;
  /** Returns false if the name was rejected (empty or duplicate). */
  onRename: (index: number, name: string) => boolean;
  onContextMenu: (index: number, x: number, y: number) => void;
  /** A tab was dropped before sheet `before` (names.length: at the end); `copy` when Ctrl was held, like Excel. */
  onMove: (from: number, before: number, copy: boolean) => void;
  /** Set by the parent to start renaming a tab (e.g. from the tab menu). */
  renaming: number | null;
  setRenaming: (index: number | null) => void;
}

/** Excel-style sheet tabs: scroll arrows, the tabs, and ⊕ to add a sheet. */
export function SheetTabs({ names, tabs, active, onSelect, onAdd, onRename, onContextMenu, onMove, renaming, setRenaming }: Props) {
  const [draft, setDraft] = useState("");
  /** While dragging a tab: which one, and the tab it would go before (names.length = the end). */
  const [drag, setDrag] = useState<{ from: number; before: number } | null>(null);
  const suppressClick = useRef(false);
  const stripRef = useRef<HTMLDivElement>(null);
  const [overflow, setOverflow] = useState({ left: false, right: false });

  function measure() {
    const el = stripRef.current;
    if (!el) return;
    const next = { left: el.scrollLeft > 0, right: el.scrollLeft + el.clientWidth < el.scrollWidth - 1 };
    setOverflow((o) => (o.left === next.left && o.right === next.right ? o : next));
  }
  const lastVisible = tabs.map((t) => !t.hidden).lastIndexOf(true);

  /** The insertion point for the pointer at `x`: before the tab under it, or after it on its right half. */
  function dropPoint(x: number): number {
    const strip = stripRef.current;
    if (!strip) return names.length;
    for (const el of strip.querySelectorAll<HTMLElement>("[data-sheet]")) {
      const box = el.getBoundingClientRect();
      if (x < box.left + box.width / 2) return Number(el.dataset.sheet);
    }
    return names.length;
  }

  /** Drag a tab to move it (Ctrl: copy). Pointer events, so focus stays in the formula being typed. */
  function startDrag(e: React.PointerEvent, from: number) {
    if (e.button !== 0) return;
    const startX = e.clientX;
    let dragging = false;
    let before = from;
    const onPointerMove = (ev: PointerEvent) => {
      if (!dragging && Math.abs(ev.clientX - startX) < 6) return;
      dragging = true;
      before = dropPoint(ev.clientX);
      setDrag({ from, before });
    };
    const onPointerUp = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
      setDrag(null);
      if (!dragging) return;
      // The click that follows a drag must not switch sheets.
      suppressClick.current = true;
      setTimeout(() => (suppressClick.current = false));
      const copy = ev.ctrlKey || ev.altKey;
      if (copy || moveTarget(from, before) !== from) onMove(from, before, copy);
    };
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
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
          tabs[i]?.hidden ? null : renaming === i ? (
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
              className={[
                i === active ? "active" : "",
                tabs[i]?.color ? "colored" : "",
                drag && drag.before === i && drag.from !== i ? "drop-before" : "",
                drag && drag.before === names.length && i === lastVisible ? "drop-after" : "",
              ]
                .filter(Boolean)
                .join(" ")}
              style={tabs[i]?.color ? ({ "--tab-color": tabs[i].color, "--tab-text": textOn(tabs[i].color!) } as React.CSSProperties) : undefined}
              title="Double-click to rename. Drag to move; Ctrl+drag to copy."
              data-sheet={i}
              onPointerDown={(e) => startDrag(e, i)}
              // Keep focus where it is: switching sheets while typing a formula points at that sheet.
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => !suppressClick.current && onSelect(i)}
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
