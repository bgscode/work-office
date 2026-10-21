import { useEffect, useRef } from "react";

export interface MenuItem {
  label: string;
  onRun: () => void;
  shortcut?: string;
  separatorBefore?: boolean;
  disabled?: boolean;
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
    <div ref={ref} className="context-menu" style={{ left, top }} role="menu">
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
            <span>{item.label}</span>
            {item.shortcut && <kbd>{item.shortcut}</kbd>}
          </button>
        </div>
      ))}
    </div>
  );
}
