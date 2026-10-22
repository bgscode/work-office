// Notes (Review → New Note, Shift+F2) and hyperlinks (Insert → Link, Ctrl+K),
// like Excel. The engine keeps them per cell (crates/sheets-engine/src/annotations.rs);
// the grid draws the red note triangle and shows notes and link tips on hover.

import { useState } from "react";
import type { RibbonCommand } from "@office/ui";
import { SheetsDocument } from "@office/engine";
import { Dialog } from "../Dialog";
import { Select } from "../Select";
import type { MenuItem } from "../ContextMenu";
import type { Rect } from "../geometry";

export interface NoteInfo {
  author: string;
  text: string;
}

export interface LinkInfo {
  target: string;
  tooltip: string;
}

export function readNote(doc: SheetsDocument, sheet: number, row: number, col: number): NoteInfo | null {
  const json = doc.noteAt(sheet, row, col);
  return json ? (JSON.parse(json) as NoteInfo) : null;
}

export function readLink(doc: SheetsDocument, sheet: number, row: number, col: number): LinkInfo | null {
  const json = doc.hyperlinkAt(sheet, row, col);
  return json ? (JSON.parse(json) as LinkInfo) : null;
}

// ---- the author of new notes

const AUTHOR_KEY = "office.userName";

/** The name notes are signed with (File → Options → User name in Excel). */
export function noteAuthor(): string {
  try {
    return localStorage.getItem(AUTHOR_KEY) ?? "";
  } catch {
    return "";
  }
}

export function setNoteAuthor(name: string) {
  try {
    localStorage.setItem(AUTHOR_KEY, name);
  } catch {
    // Private windows may refuse storage; the name is then asked again next time.
  }
}

// ---- following links

export type LinkDestination = { kind: "web"; href: string } | { kind: "place"; ref: string };

/**
 * What a link opens. Addresses typed without a scheme become web links
 * (`www.example.com`) or e-mail (`name@example.com`), as Excel does; `#Sheet2!A1`
 * is a place in the workbook. Scripts (`javascript:`) never run.
 */
export function linkDestination(target: string): LinkDestination | null {
  const t = target.trim();
  if (!t) return null;
  if (t.startsWith("#")) return { kind: "place", ref: t.slice(1) };
  if (/^(https?|mailto|ftp|file|tel):/i.test(t)) return { kind: "web", href: t };
  if (/^[a-z][a-z0-9+.-]*:/i.test(t)) return null;
  if (/^[^\s@/]+@[^\s@/]+\.[^\s@/]+$/.test(t)) return { kind: "web", href: `mailto:${t}` };
  return { kind: "web", href: `https://${t}` };
}

// ---- Insert Hyperlink

type LinkKind = "web" | "place" | "email";

interface LinkDialogProps {
  doc: SheetsDocument;
  /** The existing link, when editing one. */
  current: LinkInfo | null;
  /** The cell's text, offered as the text to display. */
  cellText: string;
  onApply: (target: string, display: string, tooltip: string) => void;
  onRemove: () => void;
  onClose: () => void;
}

/** Excel's Insert Hyperlink / Edit Hyperlink dialog (Ctrl+K). */
export function LinkDialog({ doc, current, cellText, onApply, onRemove, onClose }: LinkDialogProps) {
  const sheets = doc.sheetNames();
  const target = current?.target ?? "";
  const place = target.startsWith("#") ? target.slice(1) : "";
  const placeSheet = place.includes("!") ? place.slice(0, place.lastIndexOf("!")).replace(/^'|'$/g, "").replace(/''/g, "'") : "";
  const [kind, setKind] = useState<LinkKind>(place ? "place" : target.startsWith("mailto:") ? "email" : "web");
  const [display, setDisplay] = useState(cellText);
  const [address, setAddress] = useState(place || target.startsWith("mailto:") ? "" : target);
  const [sheetName, setSheetName] = useState(sheets.includes(placeSheet) ? placeSheet : sheets[0]);
  const [cellRef, setCellRef] = useState(place ? place.slice(place.lastIndexOf("!") + 1) : "A1");
  const [email, setEmail] = useState(target.startsWith("mailto:") ? target.slice(7).split("?")[0] : "");
  const [subject, setSubject] = useState(() => {
    const m = /[?&]subject=([^&]*)/.exec(target);
    return m ? decodeURIComponent(m[1]) : "";
  });
  const [tooltip, setTooltip] = useState(current?.tooltip ?? "");

  const built =
    kind === "web"
      ? address.trim()
      : kind === "place"
        ? `#${SheetsDocument.quoteSheetName(sheetName)}!${cellRef.trim().toUpperCase() || "A1"}`
        : email.trim()
          ? `mailto:${email.trim()}${subject ? `?subject=${encodeURIComponent(subject)}` : ""}`
          : "";
  const valid = built !== "" && (kind !== "web" || linkDestination(built) !== null);
  function ok() {
    if (!valid) return;
    onApply(built, display === cellText ? "" : display, tooltip);
    onClose();
  }
  return (
    <Dialog
      title={current ? "Edit Hyperlink" : "Insert Hyperlink"}
      onClose={onClose}
      width={460}
      footer={
        <>
          {current && (
            <button
              className="dialog-button push-left"
              onClick={() => {
                onRemove();
                onClose();
              }}
            >
              Remove Link
            </button>
          )}
          <button className="dialog-button primary" disabled={!valid} onClick={ok}>
            OK
          </button>
          <button className="dialog-button" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <div className="dialog-tabs" role="tablist" aria-label="Link to">
        {(
          [
            ["web", "Web Page or File"],
            ["place", "Place in This Document"],
            ["email", "E-mail Address"],
          ] as const
        ).map(([id, label]) => (
          <button key={id} role="tab" aria-selected={kind === id} className={kind === id ? "active" : undefined} onClick={() => setKind(id)}>
            {label}
          </button>
        ))}
      </div>
      <label className="dialog-field">
        <span>Text to display:</span>
        <input aria-label="Text to display" value={display} onChange={(e) => setDisplay(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} />
      </label>
      {kind === "web" && (
        <label className="dialog-field">
          <span>Address:</span>
          <input aria-label="Address" autoFocus placeholder="https://" value={address} onChange={(e) => setAddress(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} spellCheck={false} />
        </label>
      )}
      {kind === "place" && (
        <>
          <label className="dialog-field">
            <span>Sheet:</span>
            <Select aria-label="Sheet" value={sheetName} onChange={(e) => setSheetName(e.target.value)}>
              {sheets.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </Select>
          </label>
          <label className="dialog-field">
            <span>Cell reference:</span>
            <input aria-label="Cell reference" value={cellRef} onChange={(e) => setCellRef(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} spellCheck={false} />
          </label>
        </>
      )}
      {kind === "email" && (
        <>
          <label className="dialog-field">
            <span>E-mail address:</span>
            <input aria-label="E-mail address" autoFocus value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} spellCheck={false} />
          </label>
          <label className="dialog-field">
            <span>Subject:</span>
            <input aria-label="Subject" value={subject} onChange={(e) => setSubject(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} />
          </label>
        </>
      )}
      <label className="dialog-field">
        <span>ScreenTip:</span>
        <input aria-label="ScreenTip" value={tooltip} onChange={(e) => setTooltip(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ok()} />
      </label>
    </Dialog>
  );
}

// ---- the note editor

interface NoteEditorProps {
  note: NoteInfo | null;
  /** Where the box goes (pixels in the grid). */
  left: number;
  top: number;
  onDone: (text: string, author: string) => void;
}

/** The yellow note box, editable in place. Clicking away (or Escape) saves it, like Excel. */
export function NoteEditor({ note, left, top, onDone }: NoteEditorProps) {
  const [author, setAuthor] = useState(note?.author ?? noteAuthor());
  const [text, setText] = useState(note?.text ?? "");
  const askName = !note && !noteAuthor();
  const finish = () => {
    if (askName && author.trim()) setNoteAuthor(author.trim());
    onDone(text.replace(/\s+$/, ""), note?.author ?? author.trim());
  };
  return (
    <div className="note-popup note-editor" style={{ left, top }} onMouseDown={(e) => e.stopPropagation()}>
      {askName ? (
        <input className="note-author" aria-label="Your name" placeholder="Your name" value={author} onChange={(e) => setAuthor(e.target.value)} />
      ) : (
        author && <strong>{author}:</strong>
      )}
      <textarea
        aria-label="Note"
        autoFocus={!askName}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onBlur={(e) => {
          if (!e.currentTarget.parentElement?.contains(e.relatedTarget as Node)) finish();
        }}
        onKeyDown={(e) => {
          if (e.key === "Escape") {
            e.preventDefault();
            finish();
          }
        }}
      />
    </div>
  );
}

// ---- commands

export interface AnnotationContext {
  doc: SheetsDocument;
  sheet: number;
  row: number;
  col: number;
  /** The selection. */
  rect: Rect;
  changed: () => void;
  editNote: () => void;
  openLinkDialog: () => void;
  followLink: (row: number, col: number) => void;
  goToCell: (row: number, col: number) => void;
  showAllNotes: boolean;
  setShowAllNotes: (on: boolean) => void;
}

/** The next note after (or before) the active cell, in row order, wrapping like Excel. */
export function adjacentNote(cells: ArrayLike<number>, row: number, col: number, direction: 1 | -1): { row: number; col: number } | null {
  const list: { row: number; col: number }[] = [];
  for (let i = 0; i < cells.length; i += 2) list.push({ row: cells[i], col: cells[i + 1] });
  if (!list.length) return null;
  const after = (n: { row: number; col: number }) => n.row > row || (n.row === row && n.col > col);
  const before = (n: { row: number; col: number }) => n.row < row || (n.row === row && n.col < col);
  if (direction === 1) return list.find(after) ?? list[0];
  return [...list].reverse().find(before) ?? list[list.length - 1];
}

function deleteNotes(ctx: AnnotationContext) {
  const r = ctx.rect;
  ctx.doc.deleteNotes(ctx.sheet, r.top, r.left, r.bottom, r.right);
  ctx.changed();
}

function removeLinks(ctx: AnnotationContext, keepFormat: boolean) {
  const r = ctx.rect;
  ctx.doc.removeHyperlinks(ctx.sheet, r.top, r.left, r.bottom, r.right, keepFormat);
  ctx.changed();
}

function goToNote(ctx: AnnotationContext, direction: 1 | -1) {
  const next = adjacentNote(ctx.doc.noteCells(ctx.sheet), ctx.row, ctx.col, direction);
  if (next) ctx.goToCell(next.row, next.col);
}

/** Insert → Links. */
export function insertLinkCommand(ctx: AnnotationContext): RibbonCommand {
  return {
    id: "link",
    label: "Link",
    shortcut: "Ctrl+K",
    title: "Create a link to a web page, a place in the workbook or an e-mail address.",
    icon: "link",
    size: "large",
    group: "Links",
    onRun: ctx.openLinkDialog,
  };
}

/** Review → Notes. */
export function reviewCommands(ctx: AnnotationContext): RibbonCommand[] {
  const hasNote = readNote(ctx.doc, ctx.sheet, ctx.row, ctx.col) !== null;
  const anyNotes = ctx.doc.noteCells(ctx.sheet).length > 0;
  return [
    { id: "new-note", label: hasNote ? "Edit Note" : "New Note", shortcut: "Shift+F2", title: "Add a note about the active cell, or change its note.", icon: "note", size: "large", group: "Notes", onRun: ctx.editNote },
    { id: "delete-note", label: "Delete", title: "Delete the notes in the selected cells.", icon: "noteDelete", size: "small", group: "Notes", onRun: () => deleteNotes(ctx) },
    { id: "previous-note", label: "Previous Note", title: "Go to the previous note on this sheet.", icon: "notePrevious", size: "small", group: "Notes", onRun: () => goToNote(ctx, -1) },
    { id: "next-note", label: "Next Note", title: "Go to the next note on this sheet.", icon: "noteNext", size: "small", group: "Notes", onRun: () => goToNote(ctx, 1) },
    { id: "show-all-notes", label: "Show All Notes", title: "Show every note on the sheet.", icon: "showNotes", size: "large", group: "Notes", active: ctx.showAllNotes && anyNotes, onRun: () => ctx.setShowAllNotes(!ctx.showAllNotes) },
  ];
}

/** Right-click menu items for the active cell's note and link. */
export function annotationMenuItems(ctx: AnnotationContext): MenuItem[] {
  const hasNote = readNote(ctx.doc, ctx.sheet, ctx.row, ctx.col) !== null;
  const hasLink = readLink(ctx.doc, ctx.sheet, ctx.row, ctx.col) !== null;
  const items: MenuItem[] = [
    { label: hasLink ? "Edit Hyperlink…" : "Link…", shortcut: "Ctrl+K", icon: "link", onRun: ctx.openLinkDialog, separatorBefore: true },
  ];
  if (hasLink) {
    items.push({ label: "Open Hyperlink", onRun: () => ctx.followLink(ctx.row, ctx.col) });
    items.push({ label: "Remove Hyperlink", onRun: () => removeLinks(ctx, false) });
  }
  items.push({ label: hasNote ? "Edit Note" : "New Note", shortcut: "⇧F2", icon: "note", onRun: ctx.editNote });
  if (hasNote) items.push({ label: "Delete Note", icon: "noteDelete", onRun: () => deleteNotes(ctx) });
  return items;
}

/** Home → Clear: notes and links. */
export function clearMenuItems(ctx: AnnotationContext) {
  return [
    { id: "clear-notes", label: "Clear Notes", onRun: () => deleteNotes(ctx) },
    { id: "clear-hyperlinks", label: "Clear Hyperlinks", onRun: () => removeLinks(ctx, true) },
    { id: "remove-hyperlinks", label: "Remove Hyperlinks", onRun: () => removeLinks(ctx, false) },
  ];
}
