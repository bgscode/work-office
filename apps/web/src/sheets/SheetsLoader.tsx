// Opens a workbook before showing it: from this browser's storage or from a
// file the user picked. Big files take a moment, so an "Opening…" screen is
// painted first and the work starts on the next frame.

import { useEffect, useState } from "react";
import { SheetsDocument } from "@office/engine";
import { loadDocument, saveDocument } from "../documents";
import { Journal, journaled } from "./journal";
import { SheetsApp } from "./SheetsApp";

interface Props {
  docId: string;
  /** A file to open into this new document (from Home or File > Open). */
  initialFile?: File;
  onInitialFileOpened: () => void;
  onNewWorkbook: () => void;
  onOpenFile: (file: File) => void;
}

interface Loaded {
  /** The document, wrapped so its edits are journaled (see journal.ts). */
  doc: SheetsDocument;
  journal: Journal;
  name: string;
  /** Whether what's stored already matches the document (no autosave needed yet). */
  saved: boolean;
}

/** Let the browser paint (the "Opening…" screen) before blocking work. */
const nextFrame = () => new Promise<void>((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));

/**
 * Loads in flight, by document id. React may run the effect below twice (in
 * development) or remount the loader; both then share one load instead of
 * parsing a large file twice.
 */
const loads = new Map<string, Promise<Loaded>>();

function load(docId: string, file: File | undefined, setStatus: (s: string) => void): Promise<Loaded> {
  let promise = loads.get(docId);
  if (!promise) {
    promise = file ? openFile(file, docId, setStatus) : openSaved(docId, setStatus);
    loads.set(docId, promise);
  }
  return promise;
}

async function openFile(file: File, docId: string, setStatus: (s: string) => void): Promise<Loaded> {
  const name = file.name.replace(/\.[^.]+$/, "");
  const isCsv = file.name.toLowerCase().endsWith(".csv");
  const bytes = new Uint8Array(await file.arrayBuffer());
  setStatus(`Opening ${file.name} (${(bytes.length / 1e6).toFixed(1)} MB)…`);
  await nextFrame();
  const raw = isCsv ? SheetsDocument.fromCsv(new TextDecoder().decode(bytes)) : SheetsDocument.fromXlsx(bytes);
  // An .xlsx is stored as-is (no need to write it out again to autosave it).
  // Storing a big file takes a moment, so it happens after the grid shows.
  if (!isCsv) setTimeout(() => void saveDocument(docId, "sheets", name, bytes), 0);
  const journal = new Journal(docId);
  return { doc: journaled(raw, journal), journal, name, saved: !isCsv };
}

async function openSaved(docId: string, setStatus: (s: string) => void): Promise<Loaded> {
  const [stored, journal] = await Promise.all([loadDocument(docId), Journal.load(docId)]);
  const name = stored?.name ?? "Untitled workbook";
  if (stored) setStatus(`Opening ${name}…`);
  await nextFrame();
  let raw: SheetsDocument;
  let saved = true;
  try {
    raw = stored?.bytes ? SheetsDocument.fromXlsx(stored.bytes) : new SheetsDocument();
  } catch {
    raw = new SheetsDocument();
    saved = false;
  }
  // Edits made since the file was last fully saved.
  journal.replay(raw);
  return { doc: journaled(raw, journal), journal, name, saved };
}

export function SheetsLoader(props: Props) {
  const [loaded, setLoaded] = useState<Loaded | null>(null);
  const [status, setStatus] = useState(props.initialFile ? `Opening ${props.initialFile.name}…` : "Opening…");
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let active = true;
    const file = props.initialFile;
    if (file) props.onInitialFileOpened();
    load(props.docId, file, (s) => active && setStatus(s)).then(
      (result) => active && setLoaded(result),
      (err) => active && setError(err instanceof Error ? err.message : String(err)),
    );
    return () => {
      active = false;
    };
  }, []);

  // The document now belongs to SheetsApp; forget the finished load.
  useEffect(() => {
    if (loaded) loads.delete(props.docId);
  }, [loaded]);

  if (error) {
    return (
      <div className="sheets-loading" role="alert">
        <p>Couldn't open this file: {error}</p>
        <button onClick={props.onNewWorkbook}>Start a blank workbook</button>
      </div>
    );
  }
  if (!loaded) {
    return (
      <div className="sheets-loading" role="status" aria-live="polite">
        <span className="spinner" aria-hidden />
        <p>{status}</p>
      </div>
    );
  }
  return (
    <SheetsApp
      docId={props.docId}
      initialDoc={loaded.doc}
      journal={loaded.journal}
      initialName={loaded.name}
      initiallySaved={loaded.saved}
      onNewWorkbook={props.onNewWorkbook}
      onOpenFile={props.onOpenFile}
    />
  );
}
