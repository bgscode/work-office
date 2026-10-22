import { useEffect, useRef, useState } from "react";
import { Icon } from "@office/ui";

export interface FindOptions {
  matchCase: boolean;
  wholeCell: boolean;
}

interface Props {
  /** Open with the Replace row showing (Ctrl+H). */
  withReplace: boolean;
  /** Returns false when nothing matches. */
  onFind: (query: string, forward: boolean, opts: FindOptions) => boolean;
  /** Replace in the active cell if it matches, then find the next match. */
  onReplace: (query: string, replacement: string, opts: FindOptions) => void;
  /** Returns how many cells changed. */
  onReplaceAll: (query: string, replacement: string, opts: FindOptions) => number;
  onClose: () => void;
}

export function FindBar({ withReplace, onFind, onReplace, onReplaceAll, onClose }: Props) {
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [showReplace, setShowReplace] = useState(withReplace);
  const [opts, setOpts] = useState<FindOptions>({ matchCase: false, wholeCell: false });
  const [message, setMessage] = useState("");
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => ref.current?.focus(), []);
  useEffect(() => setShowReplace((s) => s || withReplace), [withReplace]);

  function find(forward: boolean) {
    if (query) setMessage(onFind(query, forward, opts) ? "" : "No matches");
  }

  const keys = (e: React.KeyboardEvent, onEnter: () => void) => {
    if (e.key === "Enter") {
      e.preventDefault();
      onEnter();
    } else if (e.key === "Escape") {
      e.preventDefault();
      onClose();
    }
  };

  return (
    <div className="find-bar" role="search" aria-label="Find and Replace">
      <div className="find-header">
        <span className="find-title">Find and Replace</span>
        <div className="find-tabs" role="tablist">
          <button role="tab" aria-selected={!showReplace} onClick={() => setShowReplace(false)}>
            Find
          </button>
          <button role="tab" aria-selected={showReplace} onClick={() => setShowReplace(true)}>
            Replace
          </button>
        </div>
        <button className="find-close" aria-label="Close" title="Close (Esc)" onClick={onClose}>
          <Icon name="close" size={12} />
        </button>
      </div>
      <div className="find-row">
        <input
          ref={ref}
          placeholder="Find"
          aria-label="Find"
          value={query}
          onChange={(e) => {
            setQuery(e.target.value);
            setMessage("");
          }}
          onKeyDown={(e) => keys(e, () => find(!e.shiftKey))}
        />
        <button title="Find Previous (Shift+Enter)" aria-label="Find Previous" onClick={() => find(false)}>
          <Icon name="chevronUp" size={12} />
        </button>
        <button title="Find Next (Enter)" aria-label="Find Next" onClick={() => find(true)}>
          <Icon name="chevronDown" size={12} />
        </button>
      </div>
      {showReplace && (
        <div className="find-row">
          <input
            placeholder="Replace with"
            aria-label="Replace with"
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            onKeyDown={(e) => keys(e, () => query && onReplace(query, replacement, opts))}
          />
          <button disabled={!query} onClick={() => onReplace(query, replacement, opts)}>
            Replace
          </button>
          <button
            disabled={!query}
            onClick={() => {
              const n = onReplaceAll(query, replacement, opts);
              setMessage(n === 0 ? "No matches" : `Replaced ${n} cell${n === 1 ? "" : "s"}`);
            }}
          >
            Replace All
          </button>
        </div>
      )}
      <div className="find-row options">
        <label>
          <input type="checkbox" checked={opts.matchCase} onChange={(e) => setOpts({ ...opts, matchCase: e.target.checked })} /> Match case
        </label>
        <label>
          <input type="checkbox" checked={opts.wholeCell} onChange={(e) => setOpts({ ...opts, wholeCell: e.target.checked })} /> Whole cell
        </label>
        {message && <span className="find-message">{message}</span>}
      </div>
    </div>
  );
}
