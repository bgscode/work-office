import { loadFileHandle, storeFileHandle } from "./documents";
import { desktopFileName, desktopPathOf, desktopPickWorkbook, desktopSave, desktopSaveAs, isDesktop } from "./desktop";

/** Ask the user for a file. Resolves to null if they cancel. */
export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = accept;
    input.onchange = () => resolve(input.files?.[0] ?? null);
    input.oncancel = () => resolve(null);
    input.click();
  });
}

/** File → Open / Home → Open: a workbook or text file (the native dialog in the desktop app). */
export function pickWorkbook(): Promise<File | null> {
  return isDesktop() ? desktopPickWorkbook() : pickFile(".xlsx,.csv,.tsv,.txt");
}

export function download(data: BlobPart, fileName: string, type: string) {
  if (isDesktop()) {
    // The desktop window has no downloads: ask where to save instead.
    void new Blob([data]).arrayBuffer().then((buf) => desktopSaveAs(fileName, new Uint8Array(buf)));
    return;
  }
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ---------------------------------------------------------------- saving to disk

/** The parts of the File System Access API we use (Chromium browsers). */
interface WritableFile {
  createWritable(): Promise<{ write(data: BlobPart): Promise<void>; close(): Promise<void> }>;
  name: string;
  queryPermission?(opts: { mode: "readwrite" }): Promise<PermissionState>;
  requestPermission?(opts: { mode: "readwrite" }): Promise<PermissionState>;
}
type SavePicker = (opts: { suggestedName: string; types: { description: string; accept: Record<string, string[]> }[] }) => Promise<WritableFile>;

/** Files chosen with Save As, per workbook: Save writes back to them. Also kept
 * in IndexedDB (documents.ts), so this survives a reload. */
const handles = new Map<string, WritableFile>();

/** Desktop app: the file on disk each workbook was opened from or saved to (this session). */
const desktopTargets = new Map<string, string>();

/** A workbook was opened from `file`: in the desktop app, Save writes back to it (.xlsx only). */
export function rememberOpenedFile(key: string, file: File) {
  const path = desktopPathOf(file);
  if (path && /\.xlsx$/i.test(path)) desktopTargets.set(key, path);
}

export type SaveResult = { kind: "saved"; name: string } | { kind: "downloaded" } | { kind: "cancelled" };

/** True when Save writes straight to a file on disk (no download). */
export function hasSaveTarget(key: string): boolean {
  return handles.has(key) || desktopTargets.has(key);
}

/**
 * Save As: let the user pick a file on disk (Chromium) and remember it, or
 * download a copy where the browser can't write files.
 */
export async function saveAs(key: string, data: BlobPart, fileName: string, mime: string, ext: string): Promise<SaveResult> {
  if (isDesktop()) {
    const path = await desktopSaveAs(fileName, new Uint8Array(await new Blob([data]).arrayBuffer()));
    if (!path) return { kind: "cancelled" };
    desktopTargets.set(key, path);
    return { kind: "saved", name: desktopFileName(path) };
  }
  const picker = (window as unknown as { showSaveFilePicker?: SavePicker }).showSaveFilePicker;
  if (!picker) {
    download(data, fileName, mime);
    return { kind: "downloaded" };
  }
  try {
    const handle = await picker({ suggestedName: fileName, types: [{ description: "Excel Workbook", accept: { [mime]: [ext] } }] });
    handles.set(key, handle);
    void storeFileHandle(key, handle);
    await write(handle, data);
    return { kind: "saved", name: handle.name };
  } catch (err) {
    if (err instanceof DOMException && err.name === "AbortError") return { kind: "cancelled" };
    download(data, fileName, mime);
    return { kind: "downloaded" };
  }
}

/** Save (Ctrl+S): write to the file from the last Save As, else Save As. */
export async function save(key: string, data: BlobPart, fileName: string, mime: string, ext: string): Promise<SaveResult> {
  const path = desktopTargets.get(key);
  if (path) {
    try {
      await desktopSave(path, new Uint8Array(await new Blob([data]).arrayBuffer()));
      return { kind: "saved", name: desktopFileName(path) };
    } catch {
      // Moved, deleted or read-only: ask where to save.
      desktopTargets.delete(key);
      return saveAs(key, data, fileName, mime, ext);
    }
  }
  if (isDesktop()) return saveAs(key, data, fileName, mime, ext);
  const handle = handles.get(key) ?? (await restoreHandle(key));
  if (!handle) return saveAs(key, data, fileName, mime, ext);
  try {
    await write(handle, data);
    return { kind: "saved", name: handle.name };
  } catch {
    // Permission revoked or file moved: ask again.
    handles.delete(key);
    return saveAs(key, data, fileName, mime, ext);
  }
}

/** The file from an earlier session, if the browser still lets us write to it. */
async function restoreHandle(key: string): Promise<WritableFile | null> {
  const handle = (await loadFileHandle(key)) as WritableFile | null;
  if (!handle || typeof handle.createWritable !== "function") return null;
  // After a reload the browser asks once more (Ctrl+S is the user gesture it needs).
  const opts = { mode: "readwrite" } as const;
  if ((await handle.queryPermission?.(opts)) !== "granted" && (await handle.requestPermission?.(opts)) !== "granted") return null;
  handles.set(key, handle);
  return handle;
}

async function write(handle: WritableFile, data: BlobPart) {
  const w = await handle.createWritable();
  await w.write(data);
  await w.close();
}
