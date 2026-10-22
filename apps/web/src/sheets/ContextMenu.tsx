import { useEffect, useRef, type ReactNode } from "react";
import { Icon } from "@office/ui";

export interface MenuItem {
  label: string;
  onRun: () => void;
  shortcut?: string;
  separatorBefore?: boolean;
  disabled?: boolean;
  /** Icon name from @office/ui's icon set, or any element. */
  icon?: ReactNode;
}

interface Props {
  x: number;
  y: number;
  items: MenuItem[];
  onClose: () => void;
}

/** A right-click menu at screen position (x, y). Closes on outside click or Escape. */
export function ContextMenu({ x, y, items, onClose }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    window.addEventListener("mousedown", onDown, true);
    window.addEventListener("keydown", onKey, true);
    return () => {
      window.removeEventListener("mousedown", onDown, true);
      window.removeEventListener("keydown", onKey, true);
    };
  }, [onClose]);

  // Keep the menu on screen.
  const left = Math.min(x, window.innerWidth - 240);
  const top = Math.min(y, window.innerHeight - items.length * 30 - 16);
  return (
    <div ref={ref} className="context-menu office-menu" style={{ left, top }} role="menu">
      {items.map((item) => (
        <div key={item.label}>
          {item.separatorBefore && <hr />}
          <button
            role="menuitem"
            disabled={item.disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              onClose();
              item.onRun();
            }}
          >
            <span className="menu-icon">{typeof item.icon === "string" ? <Icon name={item.icon} /> : item.icon}</span>
            <span className="menu-label">{item.label}</span>
            {item.shortcut && <kbd>{item.shortcut}</kbd>}
          </button>
        </div>
      ))}
    </div>
  );
}
