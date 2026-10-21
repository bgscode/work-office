import type { KeyboardEvent, ReactNode, RefObject } from "react";
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

  return (
    <div className="formula-bar">
      {props.nameBox}
      <span className="fx">fx</span>
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
