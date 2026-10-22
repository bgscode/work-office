// The desktop app (apps/desktop, Tauri) runs this same web app in a native
// window. There the browser's file pickers and downloads are replaced by the
// commands in apps/desktop/src-tauri/src/lib.rs: native Open/Save dialogs,
// files the OS asks us to open (double-click, Open With), and saving back to
// the file that was opened.

interface TauriGlobal {
  core: { invoke<T>(cmd: string, args?: unknown, options?: { headers: Record<string, string> }): Promise<T> };
  event: { listen(event: string, handler: () => void): Promise<() => void> };
}

function tauri(): TauriGlobal | undefined {
  return (window as unknown as { __TAURI__?: TauriGlobal }).__TAURI__;
}

export function isDesktop(): boolean {
  return tauri() !== undefined;
}

/** Where each file read from disk came from, so Save can write back to it. */
const paths = new WeakMap<File, string>();

export function desktopPathOf(file: File): string | undefined {
  return paths.get(file);
}

function fileName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}

async function readFile(path: string): Promise<File> {
  const bytes = await tauri()!.core.invoke<ArrayBuffer>("read_file", { path });
  const file = new File([bytes], fileName(path));
  paths.set(file, path);
  return file;
}

/** File → Open: the native dialog. Null when cancelled. */
export async function desktopPickWorkbook(): Promise<File | null> {
  const path = await tauri()!.core.invoke<string | null>("pick_file");
  return path ? readFile(path) : null;
}

/** Files the OS asked us to open since the last call. */
export async function takeOpenedFiles(): Promise<File[]> {
  const list = await tauri()!.core.invoke<string[]>("opened_files");
  return Promise.all(list.map(readFile));
}

/** Calls `handler` whenever the OS asks the running app to open files. */
export function onFilesOpened(handler: () => void): () => void {
  let stop: (() => void) | undefined;
  let stopped = false;
  void tauri()!
    .event.listen("files-opened", handler)
    .then((unlisten) => (stopped ? unlisten() : (stop = unlisten)));
  return () => {
    stopped = true;
    stop?.();
  };
}

/** Write to a file that was opened or saved before. */
export async function desktopSave(path: string, data: Uint8Array): Promise<void> {
  await tauri()!.core.invoke("save_file", data, { headers: { "x-path": encodeURIComponent(path) } });
}

/** Save As with the native dialog. Returns the chosen path, or null when cancelled. */
export async function desktopSaveAs(suggestedName: string, data: Uint8Array): Promise<string | null> {
  return tauri()!.core.invoke<string | null>("save_file_as", data, { headers: { "x-name": encodeURIComponent(suggestedName) } });
}

export { fileName as desktopFileName };
