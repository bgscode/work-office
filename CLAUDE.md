# Notes for AI assistants

Open-source office suite. Rust engines (`crates/`) compiled to WebAssembly,
React UI (`apps/web`). Sheets is the active app; see docs/ROADMAP.md.

## Commands

- `pnpm dev` — build wasm (debug) and run the web app
- `pnpm build:wasm` — rebuild the engine after Rust changes
- `pnpm test` / `pnpm lint` / `pnpm typecheck` — must all pass before finishing

## Rules

- Logic goes in Rust crates with unit tests; `crates/wasm` only converts types.
- Every change updates the docs it affects in the same change: `crates/sheets-engine/FUNCTIONS.md`,
  README feature list, `docs/ARCHITECTURE.md`, `docs/ROADMAP.md` checkboxes, `CHANGELOG.md` (Unreleased).
- Spreadsheet behaviour matches Excel unless documented otherwise.
- Follow CONTRIBUTING.md for how to add functions and bindings.
