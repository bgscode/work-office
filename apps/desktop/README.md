# Desktop app (Tauri)

Work Office for macOS, Windows and Linux: the web app (`apps/web`) in a native
window, with the engine running as WebAssembly inside it.

`src-tauri/src/lib.rs` adds what a browser can't do:

- registers the app for `.xlsx`, `.csv` and `.tsv` (`bundle.fileAssociations` in
  `tauri.conf.json`), so it can be the default app for them;
- opens the files the OS hands over: macOS sends `RunEvent::Opened`; Windows and
  Linux pass them as arguments, and a second launch forwards them to the running
  window (single-instance plugin);
- File → Open and Save As use the system dialogs; Save writes back to the opened
  `.xlsx` file. The page can only read or write files the user opened or picked.

The web side of this is `apps/web/src/desktop.ts`.

## Build

Install the [Tauri prerequisites](https://v2.tauri.app/start/prerequisites/) and
the repository's tools (see the main README), then from the repository root:

```sh
pnpm install
pnpm --filter @office/desktop tauri dev     # run with hot reload
pnpm --filter @office/desktop tauri build   # installer for this computer
```

The installers land in `src-tauri/target/release/bundle/`. Releases for every
system are built by `.github/workflows/release.yml` when a `v*` tag is pushed.
