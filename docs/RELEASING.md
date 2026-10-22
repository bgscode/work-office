# Releasing the desktop apps

The desktop apps are **Work Sheets**, **Work Docs** and **Work Slides**. They are
built from one Tauri project (`apps/desktop`): `tauri.conf.json` is Work Sheets,
and `tauri.docs.conf.json` / `tauri.slides.conf.json` change the name, icon,
window and update address for the other two.

## Publish a version

```sh
git tag v0.3.0
git push origin v0.3.0
```

`.github/workflows/release.yml` then:

1. creates the GitHub Release for the tag;
2. builds each app for macOS (universal: Apple silicon and Intel), Windows and
   Linux, with the version taken from the tag, and uploads the installers;
3. writes `work-sheets.json`, `work-docs.json` and `work-slides.json`: the
   files installed apps check to update themselves;
4. builds `Work.Office_<version>_universal.dmg`, one macOS disk image with all three apps.

## Auto-update

When an app starts, it reads its file from the latest release, e.g.
`https://github.com/bgscode/work-office/releases/latest/download/work-sheets.json`
(`plugins.updater` in the config). If that lists a newer version, the app
downloads it in the background and asks to restart. Updates are signed: the app
only installs files signed with the key whose public half is in `tauri.conf.json`.

The release workflow signs with two repository secrets
(GitHub → Settings → Secrets and variables → Actions → New repository secret):

| Secret | Value |
|---|---|
| `TAURI_SIGNING_PRIVATE_KEY` | the contents of the private key file (`~/.tauri/work-office.key` on the maintainer's computer) |
| `TAURI_SIGNING_PRIVATE_KEY_PASSWORD` | its password (`~/.tauri/work-office.key.password`) |

Without them the apps are still built, but with no update files, so installed
apps stay on their version. Keep the private key safe and backed up: if it is
lost, installed apps can't be updated any more (users would reinstall once with
a new key). Never commit it.

What updates itself: the macOS apps, Windows installers (`-setup.exe` and `.msi`
installs) and the Linux `.AppImage`. `.deb` and `.rpm` installs are updated by
installing the new package.

## Build on your own computer

```sh
pnpm --filter @office/desktop tauri build                                          # Work Sheets
pnpm --filter @office/desktop tauri build --config src-tauri/tauri.docs.conf.json   # Work Docs
pnpm --filter @office/desktop tauri build --config src-tauri/tauri.slides.conf.json # Work Slides
```

Local builds don't make update files (they need the private key) and aren't
signed for macOS Gatekeeper; see the README's Install section for opening them.
