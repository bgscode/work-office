// After the release builds: write work-sheets.json, work-docs.json and
// work-slides.json, the files each installed app checks to update itself
// (plugins.updater.endpoints in apps/desktop/src-tauri/tauri*.conf.json).
// Each lists, per system, the signed update file in this release and its signature.

const { GITHUB_TOKEN: token, GITHUB_REPOSITORY: repo, TAG: tag } = process.env;
const api = (path, init = {}) =>
  fetch(path.startsWith("http") ? path : `https://api.github.com/repos/${repo}${path}`, {
    ...init,
    headers: { authorization: `Bearer ${token}`, accept: "application/vnd.github+json", ...init.headers },
  });

const release = await (await api(`/releases/tags/${tag}`)).json();
const assets = release.assets ?? [];
const text = async (asset) => (await api(asset.url, { headers: { accept: "application/octet-stream" } })).text();

for (const [app, label] of [["sheets", "Sheets"], ["docs", "Docs"], ["slides", "Slides"]]) {
  const mine = assets.filter((a) => a.name.startsWith(`Work.${label}_`) || a.name.startsWith(`Work ${label}_`));
  const entry = async (suffix) => {
    const file = mine.find((a) => a.name.endsWith(suffix));
    const sig = file && mine.find((a) => a.name === `${file.name}.sig`);
    return file && sig ? { url: file.browser_download_url, signature: (await text(sig)).trim() } : null;
  };
  const mac = await entry(".app.tar.gz");
  const windows = await entry("-setup.exe");
  const linux = await entry(".AppImage");
  const platforms = Object.fromEntries(
    [
      ["darwin-aarch64", mac],
      ["darwin-x86_64", mac],
      ["windows-x86_64", windows],
      ["linux-x86_64", linux],
    ].filter(([, v]) => v),
  );
  if (Object.keys(platforms).length === 0) {
    console.log(`::warning::No signed update files for Work ${label}; skipping its update manifest.`);
    continue;
  }
  const manifest = { version: tag.replace(/^v/, ""), notes: release.body ?? "", pub_date: new Date().toISOString(), platforms };
  const name = `work-${app}.json`;
  const old = assets.find((a) => a.name === name);
  if (old) await api(`/releases/assets/${old.id}`, { method: "DELETE" });
  const upload = release.upload_url.replace(/\{.*\}$/, `?name=${name}`);
  const res = await api(upload, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(manifest, null, 2) });
  if (!res.ok) throw new Error(`Uploading ${name} failed: ${res.status} ${await res.text()}`);
  console.log(`${name}: ${Object.keys(platforms).join(", ")}`);
}
