// Keeps the open workbook in localStorage so a refresh or crash loses nothing.
// Stored as .xlsx bytes (base64), the same format as File > Download.

const KEY = "office.sheets.autosave";

interface Saved {
  name: string;
  xlsx: string;
}

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

/** Returns false if the browser refused to store it (storage full or blocked). */
export function saveDraft(name: string, xlsx: Uint8Array): boolean {
  try {
    localStorage.setItem(KEY, JSON.stringify({ name, xlsx: toBase64(xlsx) } satisfies Saved));
    return true;
  } catch {
    // Storage full or unavailable (private mode): autosave is best effort.
    return false;
  }
}

export function loadDraft(): { name: string; xlsx: Uint8Array } | null {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return null;
    const saved = JSON.parse(raw) as Saved;
    return { name: saved.name, xlsx: fromBase64(saved.xlsx) };
  } catch {
    return null;
  }
}

export function clearDraft() {
  try {
    localStorage.removeItem(KEY);
  } catch {
    // ignore
  }
}
