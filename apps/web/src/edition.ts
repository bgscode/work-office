// The desktop installs separate apps — Work Sheets, Work Docs, Work Slides —
// built from this one web app. Each opens it as `index.html?app=sheets` (see
// apps/desktop/src-tauri/tauri*.conf.json); the web version has no edition and
// shows all apps.

import { APPS } from "./chrome/AppIcon";
import type { AppKind } from "./documents";

function read(): AppKind | null {
  if (typeof window === "undefined") return null;
  const app = new URLSearchParams(window.location.search).get("app");
  return app === "sheets" || app === "docs" || app === "slides" ? app : null;
}

/** The app this window is, or null for the all-in-one web version. */
export const edition: AppKind | null = read();

/** "Work Sheets" in the Sheets app, "Work Office" on the web. */
export const productName = edition ? `Work ${APPS[edition].label}` : "Work Office";
