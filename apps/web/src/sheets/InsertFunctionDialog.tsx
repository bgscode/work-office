// Insert Function (the fx button, Shift+F3): search the function list, pick
// one, and `=NAME(` goes into the cell, like Excel's dialog.

import { useMemo, useState } from "react";
import { functionList, type FunctionInfo } from "@office/engine";
import { Dialog } from "./Dialog";
import { searchFunctions } from "./chrome";

/** Shown under "Most commonly used", in this order. */
export const COMMON_FUNCTIONS = ["SUM", "AVERAGE", "IF", "COUNT", "MAX", "MIN", "VLOOKUP", "XLOOKUP", "SUMIF", "COUNTIF", "ROUND", "IFERROR", "TODAY", "CONCAT", "INDEX", "MATCH"];

interface Props {
  onInsert: (f: FunctionInfo) => void;
  onClose: () => void;
}

export function InsertFunctionDialog({ onInsert, onClose }: Props) {
  const all = functionList();
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState<"common" | "all">("common");
  const results = useMemo(() => {
    if (query.trim()) return searchFunctions(all, query);
    if (category === "all") return all;
    return COMMON_FUNCTIONS.map((n) => all.find((f) => f.name === n)).filter((f): f is FunctionInfo => !!f);
  }, [all, query, category]);
  const [picked, setPicked] = useState<string | null>(null);
  const selected = results.find((f) => f.name === picked) ?? results[0] ?? null;

  function move(delta: number) {
    if (results.length === 0) return;
    const i = selected ? results.indexOf(selected) : -1;
    const next = results[Math.min(results.length - 1, Math.max(0, i + delta))];
    setPicked(next.name);
    document.getElementById(`fn-${next.name}`)?.scrollIntoView({ block: "nearest" });
  }

  const ok = () => selected && onInsert(selected);

  return (
    <Dialog
      title="Insert Function"
      onClose={onClose}
      width={460}
      className="insert-function"
      footer={
        <>
          <button className="dialog-button primary" disabled={!selected} onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <label className="dialog-field">
        <span>Search for a function:</span>
        <input
          autoFocus
          aria-label="Search for a function"
          placeholder="Type a short description of what you want to do"
          value={query}
          onChange={(e) => {
            setQuery(e.target.value);
            setPicked(null);
          }}
          onKeyDown={(e) => {
            if (e.key === "ArrowDown" || e.key === "ArrowUp") {
              e.preventDefault();
              move(e.key === "ArrowDown" ? 1 : -1);
            } else if (e.key === "Enter") {
              e.preventDefault();
              ok();
            }
          }}
        />
      </label>
      <label className="dialog-field inline">
        <span>Or select a category:</span>
        <select
          aria-label="Function category"
          value={query.trim() ? "search" : category}
          onChange={(e) => {
            setQuery("");
            setPicked(null);
            setCategory(e.target.value as "common" | "all");
          }}
        >
          {query.trim() && <option value="search">Recommended</option>}
          <option value="common">Most commonly used</option>
          <option value="all">All</option>
        </select>
      </label>
      <div className="dialog-field">
        <span id="fn-list-label">Select a function:</span>
        <ul
          className="function-list"
          role="listbox"
          aria-labelledby="fn-list-label"
          tabIndex={0}
          aria-activedescendant={selected ? `fn-${selected.name}` : undefined}
          onKeyDown={(e) => {
            if (e.key === "ArrowDown" || e.key === "ArrowUp") {
              e.preventDefault();
              move(e.key === "ArrowDown" ? 1 : -1);
            } else if (e.key === "Enter") {
              e.preventDefault();
              ok();
            }
          }}
        >
          {results.map((f) => (
            <li
              key={f.name}
              id={`fn-${f.name}`}
              role="option"
              aria-selected={f === selected}
              className={f === selected ? "selected" : ""}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setPicked(f.name)}
              onDoubleClick={() => onInsert(f)}
            >
              {f.name}
            </li>
          ))}
          {results.length === 0 && <li className="empty">No functions match. Try other words.</li>}
        </ul>
      </div>
      <div className="function-help" aria-live="polite">
        {selected && (
          <>
            <code>{selected.syntax}</code>
            <p>{selected.summary}</p>
          </>
        )}
      </div>
    </Dialog>
  );
}
