import { useState, type ReactNode } from "react";

export interface RibbonCommand {
  id: string;
  label: string;
  onRun?: () => void;
  /** Tooltip; defaults to the label. */
  title?: string;
  /** Shown pressed (e.g. Bold when the active cell is bold). */
  active?: boolean;
  /** Draw a divider before this command to group related commands. */
  separatorBefore?: boolean;
  /** Custom control (colour picker, dropdown) instead of a plain button. */
  render?: () => ReactNode;
}

export interface RibbonTab {
  id: string;
  label: string;
  commands: RibbonCommand[];
}

/** Office-style ribbon shared by Docs, Sheets and Slides. */
export function Ribbon({ tabs, initialTab }: { tabs: RibbonTab[]; initialTab?: string }) {
  const [active, setActive] = useState(initialTab ?? tabs[0]?.id);
  const current = tabs.find((t) => t.id === active);
  return (
    <div className="ribbon">
      <div className="ribbon-tabs" role="tablist">
        {tabs.map((t) => (
          <button key={t.id} role="tab" aria-selected={t.id === active} className={t.id === active ? "active" : ""} onClick={() => setActive(t.id)}>
            {t.label}
          </button>
        ))}
      </div>
      <div className="ribbon-commands" role="toolbar">
        {current?.commands.map((c) => (
          <span key={c.id} className="ribbon-item">
            {c.separatorBefore && <span className="ribbon-separator" aria-hidden />}
            {c.render ? (
              c.render()
            ) : (
              <button
                className={c.active ? "pressed" : ""}
                aria-pressed={c.active}
                title={c.title ?? c.label}
                // Keep focus (and the cell edit) where it is when clicking toolbar buttons.
                onMouseDown={(e) => e.preventDefault()}
                onClick={c.onRun}
              >
                {c.label}
              </button>
            )}
          </span>
        ))}
      </div>
    </div>
  );
}
