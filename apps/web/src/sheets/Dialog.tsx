// Excel-style modal dialog: title bar with a close button, body, and a footer
// with OK / Cancel. Escape cancels, Enter in a field presses the default
// button, Tab stays inside the dialog, and focus returns to where it was.
// Other features (Sort, Data Validation, Format Cells…) can reuse it.

import { useEffect, useId, useRef, type ReactNode } from "react";
import { Icon } from "@office/ui";

interface Props {
  title: string;
  onClose: () => void;
  children: ReactNode;
  /** Footer buttons; usually OK and Cancel. */
  footer?: ReactNode;
  width?: number;
  /** Extra class for the dialog box. */
  className?: string;
}

export function Dialog({ title, onClose, children, footer, width = 420, className }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const titleId = useId();
  const latestClose = useRef(onClose);
  latestClose.current = onClose;

  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const el = ref.current!;
    if (!el.contains(document.activeElement)) {
      el.querySelector<HTMLElement>("[autofocus], input, select, textarea, button:not(.dialog-close)")?.focus();
    }
    return () => previous?.focus?.({ preventScroll: true });
  }, []);

  function onKeyDown(e: React.KeyboardEvent) {
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      latestClose.current();
    } else if (e.key === "Tab") {
      // Keep focus inside the dialog.
      const focusable = [...ref.current!.querySelectorAll<HTMLElement>("button, input, select, textarea, [tabindex]:not([tabindex='-1'])")].filter(
        (x) => !(x as HTMLButtonElement).disabled,
      );
      if (focusable.length === 0) return;
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    }
  }

  return (
    <div className="dialog-backdrop" onMouseDown={(e) => e.target === e.currentTarget && e.preventDefault()}>
      <div
        ref={ref}
        className={className ? `dialog ${className}` : "dialog"}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        style={{ width }}
        onKeyDown={onKeyDown}
      >
        <div className="dialog-title">
          <span id={titleId}>{title}</span>
          <button className="dialog-close" aria-label="Close" onClick={onClose}>
            <Icon name="close" size={12} />
          </button>
        </div>
        <div className="dialog-body">{children}</div>
        {footer && <div className="dialog-footer">{footer}</div>}
      </div>
    </div>
  );
}
