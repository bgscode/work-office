// Documents saved in this browser: every open workbook is autosaved here and
// listed under "Recent" on the Home page. Each document is stored as file
// bytes (.xlsx for Sheets) in base64, plus a small index of names and times.

export type AppKind = "sheets" | "docs" | "slides";

export interface RecentDocument {
  id: string;
  app: AppKind;
  name: string;
  /** Milliseconds since 1970. */
  savedAt: number;
}

const INDEX_KEY = "office.recent";
const DOC_PREFIX = "office.doc.";
/** Older versions kept a single autosaved workbook under this key. */
const LEGACY_KEY = "office.sheets.autosave";
const MAX_DOCUMENTS = 20;

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

function fromBase64(text: string): Uint8Array {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
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
  localStorage.setItem(INDEX_KEY, JSON.stringify(list));
}

export function newDocumentId(): string {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

/** Most recently saved first. */
export function listRecent(): RecentDocument[] {
  migrateLegacyDraft();
  return readIndex().sort((a, b) => b.savedAt - a.savedAt);
}

/** Returns false if the browser refused to store it (storage full or blocked). */
export function saveDocument(id: string, app: AppKind, name: string, bytes: Uint8Array): boolean {
  try {
    localStorage.setItem(DOC_PREFIX + id, toBase64(bytes));
    const list = readIndex().filter((d) => d.id !== id);
    list.unshift({ id, app, name, savedAt: Date.now() });
    // Keep the newest documents; drop the rest to stay within storage limits.
    for (const old of list.splice(MAX_DOCUMENTS)) localStorage.removeItem(DOC_PREFIX + old.id);
    writeIndex(list);
    return true;
  } catch {
    return false;
  }
}

export function loadDocument(id: string): { name: string; bytes: Uint8Array } | null {
  try {
    const data = localStorage.getItem(DOC_PREFIX + id);
    const entry = readIndex().find((d) => d.id === id);
    if (!data || !entry) return null;
    return { name: entry.name, bytes: fromBase64(data) };
  } catch {
    return null;
  }
}

export function removeDocument(id: string) {
  try {
    localStorage.removeItem(DOC_PREFIX + id);
    writeIndex(readIndex().filter((d) => d.id !== id));
  } catch {
    // ignore
  }
}

function migrateLegacyDraft() {
  try {
    const raw = localStorage.getItem(LEGACY_KEY);
    if (!raw) return;
    const { name, xlsx } = JSON.parse(raw) as { name: string; xlsx: string };
    if (saveDocument(newDocumentId(), "sheets", name || "Untitled", fromBase64(xlsx))) localStorage.removeItem(LEGACY_KEY);
  } catch {
    // ignore a corrupt legacy draft
  }
}
