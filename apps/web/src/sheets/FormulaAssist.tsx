// Popup under the active editor: function suggestions while typing a name,
// or the syntax of the function whose arguments are being typed.

import type { FunctionInfo } from "@office/engine";

interface Props {
  /** Screen rectangle of the input being edited. */
  anchor: DOMRect | null;
  suggestions: FunctionInfo[];
  selected: number;
  /** Shown when there are no suggestions and the caret is inside a call. */
  hint: FunctionInfo | null;
  onPick: (f: FunctionInfo) => void;
  onHover: (index: number) => void;
}

export function FormulaAssist({ anchor, suggestions, selected, hint, onPick, onHover }: Props) {
  if (!anchor || (suggestions.length === 0 && !hint)) return null;
  const style = { left: anchor.left, top: anchor.bottom + 2 };
  if (suggestions.length === 0 && hint) {
    return (
      <div className="formula-hint" style={style} role="tooltip">
        <code>{hint.syntax}</code>
        <span>{hint.summary}</span>
      </div>
    );
  }
  return (
    <ul className="formula-suggestions" style={style} role="listbox" aria-label="Functions">
      {suggestions.map((f, i) => (
        <li
          key={f.name}
          role="option"
          aria-selected={i === selected}
          className={i === selected ? "selected" : ""}
          // mousedown, not click: keep focus in the editor.
          onMouseDown={(e) => {
            e.preventDefault();
            onPick(f);
          }}
          onMouseEnter={() => onHover(i)}
        >
          <code>{f.syntax}</code>
          {i === selected && <span>{f.summary}</span>}
        </li>
      ))}
    </ul>
  );
}
