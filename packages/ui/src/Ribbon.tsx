// Office-style ribbon shared by Docs, Sheets and Slides.
//
// A tab's commands are laid out in labelled groups (Excel: Clipboard, Font,
// Alignment…) in the order each group first appears. Inside a group:
//   - size "large":  tall button, 32px icon above the label (Paste, Insert Function)
//   - size "small":  16px icon + label, stacked three to a column (Cut, Copy)
//   - no size:       compact control in rows; icon only when it has an icon
//                    (Bold, Italic), else its label; `render` controls go here too.
//                    `newRow` starts a new row inside the group.
// Commands without a `group` still render, in an unlabelled group.
//
// `menu` turns a command into a dropdown (or a split button when it also has
// `onRun`). Ctrl+F1 or double-clicking a tab collapses the ribbon to its tabs;
// clicking a tab then shows the commands until one is used.

import {
  Fragment,
  useCallback,
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from "react";
import { Icon } from "./icons";

export interface RibbonMenuItem {
  id: string;
  label: string;
  onRun: () => void;
  /** Icon name from icons.tsx or any element. */
  icon?: ReactNode;
  shortcut?: string;
  disabled?: boolean;
  /** Shown with a check mark. */
  active?: boolean;
  separatorBefore?: boolean;
}

export interface RibbonCommand {
  id: string;
  label: string;
  onRun?: () => void;
  /** Longer description shown in the tooltip under the label. */
  title?: string;
  /** Shown pressed (e.g. Bold when the active cell is bold). */
  active?: boolean;
  /** Draw a divider before this command (within a row of compact commands). */
  separatorBefore?: boolean;
  /** Custom control (colour picker, dropdown) instead of a plain button. */
  render?: () => ReactNode;
  /** Ribbon group, e.g. "Clipboard", "Font", "Sort & Filter". See CONTRIBUTING.md. */
  group?: string;
  /** Icon name from icons.tsx (e.g. "bold") or any element. */
  icon?: ReactNode;
  /** "large": tall button; "small": icon + label stacked in threes; omitted: compact. */
  size?: "large" | "small";
  /** Keyboard shortcut shown in the tooltip, e.g. "Ctrl+B". */
  shortcut?: string;
  /** Dropdown items. With `onRun` too, the button is a split button. */
  menu?: RibbonMenuItem[];
  disabled?: boolean;
  /** Compact commands only: start a new row inside the group. */
  newRow?: boolean;
}

export interface RibbonTab {
  id: string;
  label: string;
  commands: RibbonCommand[];
  /** A contextual tab (e.g. Chart Design) shown only while something is selected. */
  contextual?: boolean;
}

const COLLAPSED_KEY = "office.ribbon.collapsed";

function savedCollapsed(): boolean {
  try {
    return localStorage.getItem(COLLAPSED_KEY) === "1";
  } catch {
    return false;
  }
}

function renderIcon(icon: ReactNode, size: number): ReactNode {
  return typeof icon === "string" ? <Icon name={icon} size={size} /> : icon;
}

interface Group {
  name: string;
  commands: RibbonCommand[];
}

/** Commands grouped by `group`, groups in order of first appearance. Exported for tests. */
export function groupCommands(commands: RibbonCommand[]): Group[] {
  const groups: Group[] = [];
  for (const c of commands) {
    const name = c.group ?? "";
    let g = groups.find((x) => x.name === name);
    if (!g) {
      g = { name, commands: [] };
      groups.push(g);
    }
    g.commands.push(c);
  }
  return groups;
}

type Segment =
  | { kind: "large"; command: RibbonCommand }
  | { kind: "stack"; commands: RibbonCommand[] }
  | { kind: "rows"; rows: RibbonCommand[][] };

/** How a group's commands are arranged into columns. Exported for tests. */
export function layoutGroup(commands: RibbonCommand[]): Segment[] {
  const out: Segment[] = [];
  let compact: RibbonCommand[] = [];
  const flushCompact = () => {
    if (compact.length === 0) return;
    out.push({ kind: "rows", rows: splitRows(compact) });
    compact = [];
  };
  for (const c of commands) {
    if (c.size === "large") {
      flushCompact();
      out.push({ kind: "large", command: c });
    } else if (c.size === "small") {
      flushCompact();
      const last = out[out.length - 1];
      if (last?.kind === "stack" && last.commands.length < 3) last.commands.push(c);
      else out.push({ kind: "stack", commands: [c] });
    } else {
      compact.push(c);
    }
  }
  flushCompact();
  return out;
}

function splitRows(commands: RibbonCommand[]): RibbonCommand[][] {
  if (commands.some((c) => c.newRow)) {
    const rows: RibbonCommand[][] = [];
    for (const c of commands) {
      if (c.newRow || rows.length === 0) rows.push([]);
      rows[rows.length - 1].push(c);
    }
    return rows;
  }
  const n = commands.length;
  const count = n <= 3 ? 1 : n <= 8 ? 2 : 3;
  const per = Math.ceil(n / count);
  const rows: RibbonCommand[][] = [];
  for (let i = 0; i < n; i += per) rows.push(commands.slice(i, i + per));
  return rows;
}

export function Ribbon({
  tabs,
  initialTab,
  trailing,
  showTab,
}: {
  tabs: RibbonTab[];
  initialTab?: string;
  trailing?: ReactNode;
  /** Switch to tab `id` whenever `seq` changes (e.g. Chart Design after inserting a chart). */
  showTab?: { id: string; seq: number } | null;
}) {
  const [active, setActive] = useState(initialTab ?? tabs[0]?.id);
  useEffect(() => {
    if (showTab) setActive(showTab.id);
  }, [showTab?.seq]);
  const [collapsed, setCollapsedState] = useState(savedCollapsed);
  /** While collapsed: the commands are showing temporarily. */
  const [floating, setFloating] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const tabRefs = useRef<(HTMLButtonElement | null)[]>([]);
  // A contextual tab that went away falls back to the initial tab, not File.
  const current = tabs.find((t) => t.id === active) ?? tabs.find((t) => t.id === initialTab) ?? tabs[0];
  const baseId = useId();

  const setCollapsed = useCallback((value: boolean) => {
    setCollapsedState(value);
    setFloating(false);
    try {
      localStorage.setItem(COLLAPSED_KEY, value ? "1" : "0");
    } catch {
      // ignore
    }
  }, []);

  // Ctrl+F1 collapses / expands the ribbon, like Office.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "F1" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        setCollapsedState((c) => {
          const next = !c;
          try {
            localStorage.setItem(COLLAPSED_KEY, next ? "1" : "0");
          } catch {
            // ignore
          }
          return next;
        });
        setFloating(false);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // A floating (collapsed) ribbon closes on outside click or Escape.
  useEffect(() => {
    if (!floating) return;
    const onDown = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node) && !(e.target as Element).closest?.(".ribbon-menu")) setFloating(false);
    };
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && setFloating(false);
    window.addEventListener("mousedown", onDown, true);
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("mousedown", onDown, true);
      window.removeEventListener("keydown", onKey);
    };
  }, [floating]);

  function onTabKey(e: ReactKeyboardEvent, index: number) {
    let next = -1;
    if (e.key === "ArrowRight") next = (index + 1) % tabs.length;
    else if (e.key === "ArrowLeft") next = (index - 1 + tabs.length) % tabs.length;
    else if (e.key === "Home") next = 0;
    else if (e.key === "End") next = tabs.length - 1;
    if (next < 0) return;
    e.preventDefault();
    setActive(tabs[next].id);
    tabRefs.current[next]?.focus();
  }

  const used = () => {
    if (collapsed) setFloating(false);
  };
  const showPanel = !collapsed || floating;
  const panelId = `${baseId}-panel`;

  return (
    <div ref={rootRef} className={`ribbon${collapsed ? " collapsed" : ""}`}>
      <div className="ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
        {tabs.map((t, i) => (
          <button
            key={t.id}
            ref={(el) => {
              tabRefs.current[i] = el;
            }}
            id={`${baseId}-tab-${t.id}`}
            role="tab"
            aria-selected={t.id === current?.id}
            aria-controls={panelId}
            tabIndex={t.id === current?.id ? 0 : -1}
            className={[t.id === current?.id && showPanel ? "active" : t.id === current?.id ? "current" : "", t.contextual ? "contextual" : ""].join(" ").trim()}
            data-tab={t.id}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              setActive(t.id);
              if (collapsed) setFloating((f) => !(f && t.id === current?.id));
            }}
            onDoubleClick={() => setCollapsed(!collapsed)}
            onKeyDown={(e) => onTabKey(e, i)}
          >
            {t.label}
          </button>
        ))}
        <span className="ribbon-tabs-spacer" />
        {trailing}
      </div>
      {showPanel && current && (
        <div
          id={panelId}
          className={`ribbon-panel${floating ? " floating" : ""}`}
          role="tabpanel"
          aria-labelledby={`${baseId}-tab-${current.id}`}
        >
          <div className="ribbon-commands" role="toolbar" aria-label={current.label}>
            {groupCommands(current.commands).map((g, gi) => (
              <div key={g.name || `group-${gi}`} className="ribbon-group" role="group" aria-label={g.name || undefined}>
                <div className="ribbon-group-body">
                  {layoutGroup(g.commands).map((seg, si) => (
                    <SegmentView key={si} segment={seg} onUsed={used} />
                  ))}
                </div>
                {g.name && <div className="ribbon-group-label">{g.name}</div>}
              </div>
            ))}
          </div>
          <button
            className="ribbon-collapse"
            aria-label={collapsed ? "Pin the ribbon" : "Collapse the ribbon"}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => setCollapsed(!collapsed)}
          >
            <Tooltip label={collapsed ? "Pin the ribbon" : "Collapse the ribbon"} shortcut="Ctrl+F1">
              <Icon name={collapsed ? "chevronDown" : "chevronUp"} size={12} />
            </Tooltip>
          </button>
        </div>
      )}
    </div>
  );
}

function SegmentView({ segment, onUsed }: { segment: Segment; onUsed: () => void }) {
  switch (segment.kind) {
    case "large":
      return (
        <div className="ribbon-large">
          <CommandView command={segment.command} variant="large" onUsed={onUsed} />
        </div>
      );
    case "stack":
      return (
        <div className="ribbon-stack">
          {segment.commands.map((c) => (
            <CommandView key={c.id} command={c} variant="small" onUsed={onUsed} />
          ))}
        </div>
      );
    case "rows":
      return (
        <div className="ribbon-rows">
          {segment.rows.map((row, i) => (
            <div key={i} className="ribbon-row">
              {row.map((c) => (
                <Fragment key={c.id}>
                  {c.separatorBefore && <span className="ribbon-separator" aria-hidden />}
                  <CommandView command={c} variant="compact" onUsed={onUsed} />
                </Fragment>
              ))}
            </div>
          ))}
        </div>
      );
  }
}

function CommandView({ command: c, variant, onUsed }: { command: RibbonCommand; variant: "large" | "small" | "compact"; onUsed: () => void }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const anchorRef = useRef<HTMLSpanElement>(null);
  if (c.render) {
    return (
      <span className={`ribbon-item ribbon-custom ${variant}`} data-command={c.id}>
        <Tooltip label={c.label} shortcut={c.shortcut} description={c.title}>
          {c.render()}
        </Tooltip>
      </span>
    );
  }
  const iconSize = variant === "large" ? 32 : 16;
  const hasIcon = c.icon !== undefined && c.icon !== null;
  const split = !!c.menu && !!c.onRun;
  // Large split buttons (Paste) carry the label on their lower, dropdown half.
  const showLabel = variant === "small" || (variant === "large" && !split) || (variant === "compact" && !hasIcon);
  const content = (
    <>
      {hasIcon && <span className="ribbon-icon">{renderIcon(c.icon, iconSize)}</span>}
      {showLabel && <span className="ribbon-label">{c.label}</span>}
      {c.menu && !split && <Icon name="chevronDown" size={variant === "large" ? 10 : 8} className="ribbon-chevron" />}
    </>
  );
  const run = () => {
    if (c.menu && !c.onRun) {
      setMenuOpen((o) => !o);
      return;
    }
    c.onRun?.();
    onUsed();
  };
  return (
    <span
      ref={anchorRef}
      className={`ribbon-item ${variant}${split ? " split" : ""}${c.active ? " is-pressed" : ""}`}
      data-command={c.id}
    >
      <Tooltip label={c.label} shortcut={c.shortcut} description={c.title} disabled={menuOpen}>
        <button
          className={`ribbon-button${c.active ? " pressed" : ""}`}
          aria-label={c.label}
          aria-pressed={c.active}
          aria-keyshortcuts={c.shortcut}
          aria-haspopup={c.menu && !split ? "menu" : undefined}
          aria-expanded={c.menu && !split ? menuOpen : undefined}
          disabled={c.disabled}
          // Keep focus (and the cell edit) where it is when clicking toolbar buttons.
          onMouseDown={(e) => e.preventDefault()}
          onClick={run}
        >
          {content}
        </button>
      </Tooltip>
      {split && (
        <button
          className="ribbon-split-arrow"
          aria-label={`${c.label} options`}
          aria-haspopup="menu"
          aria-expanded={menuOpen}
          disabled={c.disabled}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setMenuOpen((o) => !o)}
        >
          {variant === "large" && <span className="ribbon-label">{c.label}</span>}
          <Icon name="chevronDown" size={variant === "large" ? 10 : 8} />
        </button>
      )}
      {menuOpen && c.menu && (
        <RibbonMenu
          anchor={anchorRef.current}
          items={c.menu}
          label={c.label}
          onClose={() => setMenuOpen(false)}
          onRun={(item) => {
            setMenuOpen(false);
            item.onRun();
            onUsed();
          }}
        />
      )}
    </span>
  );
}

function RibbonMenu({
  anchor,
  items,
  label,
  onClose,
  onRun,
}: {
  anchor: HTMLElement | null;
  items: RibbonMenuItem[];
  label: string;
  onClose: () => void;
  onRun: (item: RibbonMenuItem) => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    const r = anchor?.getBoundingClientRect();
    const menu = ref.current;
    if (!r || !menu) return;
    const left = Math.max(4, Math.min(r.left, window.innerWidth - menu.offsetWidth - 4));
    setPos({ left, top: r.bottom + 2 });
    menu.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus({ preventScroll: true });
  }, [anchor]);

  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node) && !anchor?.contains(e.target as Node)) onClose();
    };
    window.addEventListener("mousedown", onDown, true);
    return () => window.removeEventListener("mousedown", onDown, true);
  }, [anchor, onClose]);

  function onKey(e: ReactKeyboardEvent) {
    const buttons = [...(ref.current?.querySelectorAll<HTMLButtonElement>("button:not(:disabled)") ?? [])];
    const i = buttons.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const next = e.key === "ArrowDown" ? (i + 1) % buttons.length : (i - 1 + buttons.length) % buttons.length;
      buttons[next]?.focus();
    } else if (e.key === "Escape" || e.key === "Tab") {
      e.preventDefault();
      e.stopPropagation();
      onClose();
      anchor?.querySelector("button")?.focus();
    }
  }

  return (
    <div
      ref={ref}
      className="ribbon-menu office-menu"
      role="menu"
      aria-label={label}
      style={pos ? { left: pos.left, top: pos.top } : { visibility: "hidden" }}
      onKeyDown={onKey}
    >
      {items.map((item) => (
        <Fragment key={item.id}>
          {item.separatorBefore && <hr />}
          <button
            role="menuitem"
            data-command={item.id}
            disabled={item.disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onRun(item)}
          >
            <span className="menu-icon">{item.active ? <Icon name="check" /> : renderIcon(item.icon, 16)}</span>
            <span className="menu-label">{item.label}</span>
            {item.shortcut && <kbd>{item.shortcut}</kbd>}
          </button>
        </Fragment>
      ))}
    </div>
  );
}

/**
 * Office-style tooltip: bold title with the shortcut, then a description.
 * Shown after hovering for a moment; never steals focus.
 */
export function Tooltip({
  label,
  shortcut,
  description,
  disabled,
  children,
}: {
  label: string;
  shortcut?: string;
  description?: string;
  disabled?: boolean;
  children: ReactNode;
}) {
  const [at, setAt] = useState<{ x: number; y: number } | null>(null);
  const timer = useRef<number>();
  const wrapRef = useRef<HTMLSpanElement>(null);
  useEffect(() => () => window.clearTimeout(timer.current), []);
  useEffect(() => {
    if (disabled) setAt(null);
  }, [disabled]);
  const show = () => {
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => {
      const r = wrapRef.current?.getBoundingClientRect();
      if (r) setAt({ x: Math.min(r.left, window.innerWidth - 280), y: r.bottom + 6 });
    }, 550);
  };
  const hide = () => {
    window.clearTimeout(timer.current);
    setAt(null);
  };
  const extra = description && description !== label ? description : undefined;
  return (
    <span ref={wrapRef} className="tooltip-anchor" onMouseEnter={show} onMouseLeave={hide} onMouseDown={hide}>
      {children}
      {at && !disabled && (
        <span className="office-tooltip" role="tooltip" style={{ left: at.x, top: at.y }}>
          <strong>
            {label}
            {shortcut && ` (${shortcut})`}
          </strong>
          {extra && <span>{extra}</span>}
        </span>
      )}
    </span>
  );
}
