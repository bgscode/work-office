import type { KeyboardEvent, MouseEvent, ReactNode, RefObject } from "react";
import { Icon, Tooltip } from "@office/ui";
import { FORMULA_BAR_ID, type EditState } from "./types";

interface Props {
  inputRef: RefObject<HTMLTextAreaElement>;
  /** The Name Box shown on the left. */
  nameBox: ReactNode;
  /** The active cell's stored input, shown when not editing. */
  input: string;
  editing: EditState | null;
  onStartEdit: () => void;
  onDraftChange: (draft: string, caret: number) => void;
  onCaret: (caret: number) => void;
  /** Gets first look at keys (autocomplete); returns true if it handled the key. */
  onEditorKey: (e: KeyboardEvent<HTMLTextAreaElement>) => boolean;
  onCommit: (dRow: number, dCol: number) => void;
  onCancel: () => void;
  /** The fx button: open the Insert Function dialog. */
  onInsertFunction: () => void;
}

export function FormulaBar(props: Props) {
  const { inputRef, input, editing } = props;
  const value = editing ? editing.draft : input;

  function onKeyDown(e: KeyboardEvent<HTMLTextAreaElement>) {
    if (e.nativeEvent.isComposing || props.onEditorKey(e)) return;
    if (e.key === "Enter") {
      e.preventDefault();
      props.onCommit(e.shiftKey ? -1 : 1, 0);
    } else if (e.key === "Tab") {
      e.preventDefault();
      props.onCommit(0, e.shiftKey ? -1 : 1);
    } else if (e.key === "Escape") {
      e.preventDefault();
      props.onCancel();
    }
  }

  // Clicking the bar's buttons must not end the edit by moving focus.
  const keepFocus = (e: MouseEvent) => e.preventDefault();

  return (
    <div className="formula-bar">
      {props.nameBox}
      <span className="formula-bar-divider" aria-hidden />
      <span className="formula-buttons">
        <Tooltip label="Cancel" shortcut="Esc">
          <button className="fb-cancel" aria-label="Cancel" disabled={!editing} onMouseDown={keepFocus} onClick={props.onCancel}>
            <Icon name="close" size={14} />
          </button>
        </Tooltip>
        <Tooltip label="Enter" shortcut="Enter">
          <button className="fb-enter" aria-label="Enter" disabled={!editing} onMouseDown={keepFocus} onClick={() => props.onCommit(0, 0)}>
            <Icon name="check" size={14} />
          </button>
        </Tooltip>
        <Tooltip label="Insert Function" shortcut="Shift+F3">
          <button className="fx" aria-label="Insert Function" onMouseDown={keepFocus} onClick={props.onInsertFunction}>
            <Icon name="function" size={16} />
          </button>
        </Tooltip>
      </span>
      {/* The wrapper keeps the bar one line tall; while editing a multi-line
          cell the textarea grows downward over the grid instead of pushing it. */}
      <div className="formula-input">
        <textarea
          ref={inputRef}
          id={FORMULA_BAR_ID}
          aria-label="Formula"
          spellCheck={false}
          autoComplete="off"
          rows={editing?.source === "bar" ? Math.min(6, value.split("\n").length) : 1}
          value={value}
          onFocus={() => {
            if (!editing) props.onStartEdit();
          }}
          onChange={(e) => props.onDraftChange(e.target.value, e.target.selectionStart ?? e.target.value.length)}
          onSelect={(e) => editing && props.onCaret(e.currentTarget.selectionStart ?? 0)}
          onKeyDown={onKeyDown}
          onBlur={(e) => {
            if (editing?.source === "bar" && !e.relatedTarget?.classList.contains("cell-editor")) props.onCommit(0, 0);
          }}
        />
      </div>
    </div>
  );
}
