import { useState } from "react";

interface Props {
  names: string[];
  active: number;
  onSelect: (index: number) => void;
  onAdd: () => void;
  /** Returns false if the name was rejected (empty or duplicate). */
  onRename: (index: number, name: string) => boolean;
  onContextMenu: (index: number, x: number, y: number) => void;
  /** Set by the parent to start renaming a tab (e.g. from the tab menu). */
  renaming: number | null;
  setRenaming: (index: number | null) => void;
}

export function SheetTabs({ names, active, onSelect, onAdd, onRename, onContextMenu, renaming, setRenaming }: Props) {
  const [draft, setDraft] = useState("");
  return (
    <div className="sheet-tabs" role="tablist">
      <button className="add-sheet" onClick={onAdd} title="Add sheet" aria-label="Add sheet">
        +
      </button>
      {names.map((name, i) =>
        renaming === i ? (
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
            className={i === active ? "active" : ""}
            title="Double-click to rename"
            onClick={() => onSelect(i)}
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
  );
}
