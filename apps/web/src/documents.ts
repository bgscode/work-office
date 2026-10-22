// Documents saved in this browser: every open workbook is autosaved here and
// listed under "Recent" on the Home page.
//
// File bytes (.xlsx for Sheets) live in IndexedDB, which holds hundreds of
// megabytes and stores binary data directly. A small index of names and times
// stays in localStorage so Home can list recent files instantly.

export type AppKind = "sheets" | "docs" | "slides";

export interface RecentDocument {
  id: string;
  app: AppKind;
  name: string;
  /** Milliseconds since 1970. */
  savedAt: number;
}

const INDEX_KEY = "office.recent";
/** Older versions stored base64 bytes in localStorage under these keys. */
const LEGACY_DOC_PREFIX = "office.doc.";
const LEGACY_DRAFT_KEY = "office.sheets.autosave";
const MAX_DOCUMENTS = 20;

const DB_NAME = "office";
const STORE = "documents";

let dbPromise: Promise<IDBDatabase> | null = null;
function db(): Promise<IDBDatabase> {
  dbPromise ??= new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

async function idb<T>(mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const database = await db();
  return new Promise((resolve, reject) => {
    const req = run(database.transaction(STORE, mode).objectStore(STORE));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function readIndex(): RecentDocument[] {
  try {
    const list = JSON.parse(localStorage.getItem(INDEX_KEY) ?? "[]") as RecentDocument[];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

function writeIndex(list: RecentDocument[]) {
  try {
    localStorage.setItem(INDEX_KEY, JSON.stringify(list));
  } catch {
    // ignore: the list is a convenience
  }
}

function fromBase64(text: string): Uint8Array {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function newDocumentId(): string {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

/** Most recently saved first. */
export function listRecent(): RecentDocument[] {
  void migrateLegacyDraft();
  return readIndex().sort((a, b) => b.savedAt - a.savedAt);
}

/** Record a document in Recent without touching its bytes (e.g. after a rename). */
export function touchDocument(id: string, app: AppKind, name: string) {
  const list = readIndex().filter((d) => d.id !== id);
  list.unshift({ id, app, name, savedAt: Date.now() });
  const dropped = list.splice(MAX_DOCUMENTS);
  writeIndex(list);
  for (const old of dropped) void idb("readwrite", (s) => s.delete(old.id)).catch(() => {});
}

/** Store a document's bytes. Resolves to false if the browser refused (storage full or blocked). */
export async function saveDocument(id: string, app: AppKind, name: string, bytes: Uint8Array): Promise<boolean> {
  try {
    await idb("readwrite", (s) => s.put(bytes, id));
    touchDocument(id, app, name);
    return true;
  } catch {
    return false;
  }
}

/**
 * A stored document. `bytes` is null for a document that only has a journal
 * so far (a new workbook that was edited but never fully saved).
 */
export async function loadDocument(id: string): Promise<{ name: string; bytes: Uint8Array | null } | null> {
  const entry = readIndex().find((d) => d.id === id);
  if (!entry) return null;
  try {
    const bytes = await idb<Uint8Array | undefined>("readonly", (s) => s.get(id));
    if (bytes) return { name: entry.name, bytes };
  } catch {
    // fall through to the legacy store
  }
  // Saved by an older version: base64 in localStorage. Move it to IndexedDB.
  const legacy = localStorage.getItem(LEGACY_DOC_PREFIX + id);
  if (!legacy) return { name: entry.name, bytes: null };
  const bytes = fromBase64(legacy);
  if (await saveDocument(id, entry.app, entry.name, bytes)) localStorage.removeItem(LEGACY_DOC_PREFIX + id);
  return { name: entry.name, bytes };
}

export async function removeDocument(id: string) {
  writeIndex(readIndex().filter((d) => d.id !== id));
  localStorage.removeItem(LEGACY_DOC_PREFIX + id);
  try {
    await idb("readwrite", (s) => s.delete(id));
    await idb("readwrite", (s) => s.delete(journalKey(id)));
  } catch {
    // ignore
  }
}

// ---- edit journal (see sheets/journal.ts)

const journalKey = (id: string) => `${id}#journal`;

/** Replace the stored journal of a document. */
export async function writeJournal(id: string, entries: unknown[]): Promise<boolean> {
  try {
    await idb("readwrite", (s) => s.put(entries, journalKey(id)));
    return true;
  } catch {
    return false;
  }
}

export async function readJournal(id: string): Promise<unknown[]> {
  try {
    return (await idb<unknown[] | undefined>("readonly", (s) => s.get(journalKey(id)))) ?? [];
  } catch {
    return [];
  }
}

async function migrateLegacyDraft() {
  try {
    const raw = localStorage.getItem(LEGACY_DRAFT_KEY);
    if (!raw) return;
    localStorage.removeItem(LEGACY_DRAFT_KEY);
    const { name, xlsx } = JSON.parse(raw) as { name: string; xlsx: string };
    await saveDocument(newDocumentId(), "sheets", name || "Untitled", fromBase64(xlsx));
  } catch {
    // ignore a corrupt legacy draft
  }
}
