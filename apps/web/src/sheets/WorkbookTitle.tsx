// The workbook's part of the title bar: quick-access Save/Undo/Redo, the file
// name (click to rename) and the autosave status, like Excel's "Book1 • Saved".
// Rendered into App's title bar through a portal.

import { useContext, useState } from "react";
import { createPortal } from "react-dom";
import { Icon, Tooltip } from "@office/ui";
import { TitleSlot } from "../chrome/titleSlot";
import { MOD_KEY } from "./chrome";

export type SaveStatus = "saved" | "saving" | "failed";

interface Props {
  name: string;
  /** Rename the workbook; the name is already cleaned and non-empty. */
  onRename: (name: string) => void;
  /** Called when renaming ends, to give focus back to the grid. */
  onDone: () => void;
  status: SaveStatus;
  onSave: () => void;
  onUndo: () => void;
  onRedo: () => void;
}

/** Characters Windows and Excel don't allow in file names. */
export function cleanFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 200);
}

const STATUS_TEXT: Record<SaveStatus, string> = {
  saved: "Saved to this browser",
  saving: "Saving…",
  failed: "Not saved: browser storage is full or blocked",
};

export function WorkbookTitle(props: Props) {
  const slot = useContext(TitleSlot);
  const [draft, setDraft] = useState<string | null>(null);
  if (!slot) return null;

  function finish(commit: boolean) {
    if (commit && draft !== null) {
      const name = cleanFileName(draft);
      if (name && name !== props.name) props.onRename(name);
    }
    setDraft(null);
    props.onDone();
  }

  return createPortal(
    <div className="workbook-title">
      <div className="quick-access" role="toolbar" aria-label="Quick Access Toolbar">
        <Tooltip label="Save a copy (.xlsx)" description="Download this workbook as an Excel file.">
          <button aria-label="Save a copy" onMouseDown={(e) => e.preventDefault()} onClick={props.onSave}>
            <Icon name="save" />
          </button>
        </Tooltip>
        <Tooltip label="Undo" shortcut={`${MOD_KEY}Z`}>
          <button aria-label="Undo" onMouseDown={(e) => e.preventDefault()} onClick={props.onUndo}>
            <Icon name="undo" />
          </button>
        </Tooltip>
        <Tooltip label="Redo" shortcut={`${MOD_KEY}Y`}>
          <button aria-label="Redo" onMouseDown={(e) => e.preventDefault()} onClick={props.onRedo}>
            <Icon name="redo" />
          </button>
        </Tooltip>
      </div>
      <div className="file-title">
        {draft === null ? (
          <button
            className="file-name"
            aria-label={`File name: ${props.name}. Click to rename`}
            title="Rename this workbook"
            onClick={() => setDraft(props.name)}
          >
            {props.name}
          </button>
        ) : (
          <input
            className="file-name-input"
            aria-label="File name"
            autoFocus
            value={draft}
            maxLength={200}
            onFocus={(e) => e.target.select()}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                finish(true);
              } else if (e.key === "Escape") {
                e.preventDefault();
                finish(false);
              }
            }}
            onBlur={() => finish(true)}
          />
        )}
        <span className={`save-status ${props.status}`} role="status" title={STATUS_TEXT[props.status]}>
          {props.status === "saving" ? "Saving…" : props.status === "saved" ? "Saved" : "Not saved"}
        </span>
      </div>
    </div>,
    slot,
  );
}
