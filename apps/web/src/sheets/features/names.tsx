// Defined names: Formulas → Name Manager (Ctrl+F3) and Define Name, like
// Excel's dialogs. Names are kept by the engine (crates/sheets-engine/src/names.rs);
// a name refers to a range (`=Sheet1!$B$2:$B$100`) or a value or formula (`=0.18`).

import { useMemo, useState } from "react";
import { SheetsDocument } from "@office/engine";
import { Dialog } from "../Dialog";
import { Select } from "../Select";
import { colName, type Rect } from "../geometry";

export interface NameInfo {
  name: string;
  refersTo: string;
  /** Sheet index, or -1 for the whole workbook. */
  scope: number;
  comment: string;
}

export function readNames(doc: SheetsDocument): NameInfo[] {
  return JSON.parse(doc.namesJson()) as NameInfo[];
}

/** `Sheet1!$A$1:$B$5` for a selection, as Define Name fills in. */
export function absoluteRef(sheetName: string, r: Rect): string {
  const cell = (row: number, col: number) => `$${colName(col)}$${row + 1}`;
  const area = r.top === r.bottom && r.left === r.right ? cell(r.top, r.left) : `${cell(r.top, r.left)}:${cell(r.bottom, r.right)}`;
  return `${SheetsDocument.quoteSheetName(sheetName)}!${area}`;
}

interface EditProps {
  doc: SheetsDocument;
  /** The name being edited, or null for a new one. */
  editing: NameInfo | null;
  /** Initial "Refers to" for a new name (the selection). */
  initialRef: string;
  /** Suggested name for a new one (e.g. the header above the selection). */
  initialName?: string;
  onDone: () => void;
  onClose: () => void;
}

/** New Name / Edit Name. */
export function NameEditDialog({ doc, editing, initialRef, initialName = "", onDone, onClose }: EditProps) {
  const sheets = doc.sheetNames();
  const [name, setName] = useState(editing?.name ?? initialName);
  const [scope, setScope] = useState(String(editing?.scope ?? -1));
  const [comment, setComment] = useState(editing?.comment ?? "");
  const [refersTo, setRefersTo] = useState("=" + (editing?.refersTo ?? initialRef));
  const [error, setError] = useState("");
  function ok() {
    const err = doc.defineName(name.trim(), refersTo, Number(scope), comment, editing?.name ?? "", editing?.scope ?? -1);
    if (err) {
      setError(err);
      return;
    }
    onDone();
    onClose();
  }
  return (
    <Dialog
      title={editing ? "Edit Name" : "New Name"}
      onClose={onClose}
      width={440}
      footer={
        <>
          <button className="dialog-button primary" onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <label className="dialog-field">
        <span>Name:</span>
        <input aria-label="Name" autoFocus value={name} onChange={(e) => (setName(e.target.value), setError(""))} onKeyDown={(e) => e.key === "Enter" && ok()} />
      </label>
      <label className="dialog-field">
        <span>Scope:</span>
        <Select aria-label="Scope" value={scope} onChange={(e) => setScope(e.target.value)}>
          <option value="-1">Workbook</option>
          {sheets.map((s, i) => (
            <option key={i} value={String(i)}>
              {s}
            </option>
          ))}
        </Select>
      </label>
      <label className="dialog-field">
        <span>Comment:</span>
        <textarea aria-label="Comment" rows={2} value={comment} onChange={(e) => setComment(e.target.value)} />
      </label>
      <label className="dialog-field">
        <span>Refers to:</span>
        <input aria-label="Refers to" value={refersTo} onChange={(e) => (setRefersTo(e.target.value), setError(""))} onKeyDown={(e) => e.key === "Enter" && ok()} spellCheck={false} />
      </label>
      {error && <p className="dialog-error">{error}</p>}
    </Dialog>
  );
}

interface ManagerProps {
  doc: SheetsDocument;
  revision: number;
  /** For New: the selection as a reference. */
  selectionRef: string;
  onChanged: () => void;
  onClose: () => void;
}

/** Name Manager: every name with what it refers to, its current value and scope. */
export function NameManagerDialog({ doc, revision, selectionRef, onChanged, onClose }: ManagerProps) {
  const names = useMemo(() => readNames(doc), [doc, revision]);
  const sheets = doc.sheetNames();
  const [current, setCurrent] = useState(0);
  const [edit, setEdit] = useState<{ name: NameInfo | null } | null>(null);
  const [filter, setFilter] = useState("");
  const shown = names.filter((n) => !filter || n.name.toLowerCase().includes(filter.toLowerCase()));
  const selected = shown[Math.min(current, shown.length - 1)];
  return (
    <>
      <Dialog
        title="Name Manager"
        onClose={onClose}
        width={640}
        footer={
          <button className="dialog-button" onClick={onClose}>
            Close
          </button>
        }
      >
        <div className="dialog-row toolbar">
          <button className="dialog-button" onClick={() => setEdit({ name: null })}>
            New…
          </button>
          <button className="dialog-button" disabled={!selected} onClick={() => selected && setEdit({ name: selected })}>
            Edit…
          </button>
          <button
            className="dialog-button"
            disabled={!selected}
            onClick={() => {
              if (selected && confirm(`Delete the name "${selected.name}"? Formulas using it will show #NAME?. You can undo this.`)) {
                doc.deleteName(selected.name, selected.scope);
                onChanged();
              }
            }}
          >
            Delete
          </button>
          <input className="push" aria-label="Filter names" placeholder="Filter" value={filter} onChange={(e) => setFilter(e.target.value)} />
        </div>
        {shown.length === 0 ? (
          <p className="dialog-message">
            {names.length === 0
              ? "No names yet. A name like Sales or GST makes formulas easier to read: =SUM(Sales)*(1+GST). Click New… or type a name in the Name Box."
              : "No names match."}
          </p>
        ) : (
          <table className="dialog-table names-table">
            <thead>
              <tr>
                <th>Name</th>
                <th>Refers To</th>
                <th>Scope</th>
                <th>Comment</th>
              </tr>
            </thead>
            <tbody>
              {shown.map((n, i) => (
                <tr key={`${n.name}:${n.scope}`} className={n === selected ? "current" : undefined} onMouseDown={() => setCurrent(i)} onDoubleClick={() => setEdit({ name: n })}>
                  <td>{n.name}</td>
                  <td>={n.refersTo}</td>
                  <td>{n.scope < 0 ? "Workbook" : sheets[n.scope]}</td>
                  <td>{n.comment}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Dialog>
      {edit && <NameEditDialog doc={doc} editing={edit.name} initialRef={selectionRef} onDone={onChanged} onClose={() => setEdit(null)} />}
    </>
  );
}
