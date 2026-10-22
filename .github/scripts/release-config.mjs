// Writes apps/desktop/src-tauri/tauri.release.conf.json for a release build:
// the version from the tag (v1.2.3 -> 1.2.3) and, when the signing key secret is
// set, the signed update files the apps download to update themselves.
import { writeFileSync } from "node:fs";

const tag = process.env.TAG ?? "";
const version = tag.replace(/^v/, "");
if (!/^\d+\.\d+\.\d+$/.test(version)) throw new Error(`Release tags look like v1.2.3, not "${tag}"`);
const signed = Boolean(process.env.TAURI_SIGNING_PRIVATE_KEY);
if (!signed) console.log("::warning::TAURI_SIGNING_PRIVATE_KEY is not set: the apps will be built without auto-update files.");
writeFileSync(
  new URL("../../apps/desktop/src-tauri/tauri.release.conf.json", import.meta.url),
  JSON.stringify({ version, bundle: { createUpdaterArtifacts: signed } }, null, 2),
);
