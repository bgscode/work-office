//! The desktop app: the web app (apps/web) in a native window, plus what a
//! browser can't do — be the default app for .xlsx/.csv files, open the file
//! that was double-clicked, and save back to it with native dialogs.
//!
//! The web page can only read and write files the user chose: ones the OS
//! opened with this app, or picked in an Open/Save dialog (`Files::allowed`).

use std::collections::HashSet;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{AppHandle, Emitter, Manager, State};
use tauri_plugin_dialog::DialogExt;

const WORKBOOK_EXTENSIONS: &[&str] = &["xlsx", "csv", "tsv", "txt"];

#[derive(Default)]
struct Files {
    /// Opened by the OS, not yet taken by the page.
    pending: Mutex<Vec<PathBuf>>,
    /// Files the page may read or write.
    allowed: Mutex<HashSet<PathBuf>>,
}

impl Files {
    fn allow(&self, path: &Path) {
        self.allowed.lock().unwrap().insert(path.to_path_buf());
    }

    fn check(&self, path: &Path) -> Result<(), String> {
        if self.allowed.lock().unwrap().contains(path) {
            Ok(())
        } else {
            Err(format!("{} wasn't opened in this app", path.display()))
        }
    }
}

/// The OS asked us to open files: queue them and tell the page.
fn open_paths(app: &AppHandle, paths: impl IntoIterator<Item = PathBuf>) {
    let files = app.state::<Files>();
    let mut added = false;
    for path in paths {
        let is_workbook = path
            .extension()
            .and_then(|e| e.to_str())
            .is_some_and(|e| WORKBOOK_EXTENSIONS.iter().any(|w| w.eq_ignore_ascii_case(e)));
        if is_workbook && path.is_file() {
            files.allow(&path);
            files.pending.lock().unwrap().push(path);
            added = true;
        }
    }
    if added {
        let _ = app.emit("files-opened", ());
        if let Some(window) = app.get_webview_window("main") {
            let _ = window.unminimize();
            let _ = window.set_focus();
        }
    }
}

/// File paths among command-line arguments (Windows and Linux pass the file to open this way).
fn paths_in_args(args: &[String], cwd: &Path) -> Vec<PathBuf> {
    args.iter()
        .skip(1)
        .filter(|a| !a.starts_with('-'))
        .map(|a| {
            let p = PathBuf::from(a);
            if p.is_absolute() {
                p
            } else {
                cwd.join(p)
            }
        })
        .collect()
}

#[tauri::command]
fn opened_files(files: State<Files>) -> Vec<String> {
    files.pending.lock().unwrap().drain(..).map(|p| p.to_string_lossy().into_owned()).collect()
}

/// File → Open.
#[tauri::command]
async fn pick_file(app: AppHandle, files: State<'_, Files>) -> Result<Option<String>, String> {
    let picked = app
        .dialog()
        .file()
        .add_filter("Workbooks and text files", WORKBOOK_EXTENSIONS)
        .add_filter("Excel Workbook", &["xlsx"])
        .add_filter("CSV", &["csv"])
        .blocking_pick_file();
    let Some(path) = picked.and_then(|p| p.into_path().ok()) else { return Ok(None) };
    files.allow(&path);
    Ok(Some(path.to_string_lossy().into_owned()))
}

#[tauri::command]
fn read_file(path: String, files: State<Files>) -> Result<Response, String> {
    let path = PathBuf::from(path);
    files.check(&path)?;
    std::fs::read(&path).map(Response::new).map_err(|e| e.to_string())
}

/// Save: write the request body to the file in the `x-path` header.
#[tauri::command]
fn save_file(request: Request, files: State<Files>) -> Result<(), String> {
    let path = PathBuf::from(header(&request, "x-path")?);
    files.check(&path)?;
    write_atomic(&path, body(&request)?)
}

/// Save As: ask where, write the body there, and return the path.
#[tauri::command]
async fn save_file_as(app: AppHandle, request: Request<'_>, files: State<'_, Files>) -> Result<Option<String>, String> {
    let name = header(&request, "x-name")?;
    let ext = Path::new(&name).extension().and_then(|e| e.to_str()).unwrap_or("xlsx").to_ascii_lowercase();
    let label = match ext.as_str() {
        "xlsx" => "Excel Workbook",
        "csv" => "CSV (Comma delimited)",
        "pdf" => "PDF",
        _ => "File",
    };
    let chosen = app.dialog().file().set_file_name(&name).add_filter(label, &[ext.as_str()]).blocking_save_file();
    let Some(path) = chosen.and_then(|p| p.into_path().ok()) else { return Ok(None) };
    write_atomic(&path, body(&request)?)?;
    files.allow(&path);
    Ok(Some(path.to_string_lossy().into_owned()))
}

fn header(request: &Request, name: &str) -> Result<String, String> {
    let value = request.headers().get(name).ok_or(format!("missing {name}"))?;
    percent_decode(value.to_str().map_err(|e| e.to_string())?)
}

fn body<'a>(request: &'a Request) -> Result<&'a [u8], String> {
    match request.body() {
        InvokeBody::Raw(bytes) => Ok(bytes),
        InvokeBody::Json(_) => Err("expected the file's bytes".into()),
    }
}

/// Write next to the file, then replace it, so a failed save never leaves half a file.
fn write_atomic(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let name = path.file_name().ok_or("not a file path")?.to_string_lossy();
    let temp = path.with_file_name(format!(".{name}.saving"));
    std::fs::write(&temp, bytes).map_err(|e| e.to_string())?;
    std::fs::rename(&temp, path).map_err(|e| {
        let _ = std::fs::remove_file(&temp);
        e.to_string()
    })
}

/// Undo JavaScript's `encodeURIComponent` (headers must be ASCII).
fn percent_decode(s: &str) -> Result<String, String> {
    let b = s.as_bytes();
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%' && i + 2 < b.len() {
            let hex = std::str::from_utf8(&b[i + 1..i + 3]).map_err(|e| e.to_string())?;
            out.push(u8::from_str_radix(hex, 16).map_err(|e| e.to_string())?);
            i += 3;
        } else {
            out.push(b[i]);
            i += 1;
        }
    }
    String::from_utf8(out).map_err(|e| e.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let app = tauri::Builder::default()
        // A second launch (double-clicking another file on Windows/Linux) opens it in this window.
        .plugin(tauri_plugin_single_instance::init(|app, args, cwd| {
            open_paths(app, paths_in_args(&args, Path::new(&cwd)));
        }))
        .plugin(tauri_plugin_dialog::init())
        .manage(Files::default())
        .invoke_handler(tauri::generate_handler![opened_files, pick_file, read_file, save_file, save_file_as])
        .setup(|app| {
            let args: Vec<String> = std::env::args().collect();
            let cwd = std::env::current_dir().unwrap_or_default();
            open_paths(app.handle(), paths_in_args(&args, &cwd));
            Ok(())
        })
        .build(tauri::generate_context!())
        .expect("error while starting Work Office");

    app.run(|_app, _event| {
        // macOS hands over double-clicked files as an event, also at launch.
        #[cfg(any(target_os = "macos", target_os = "ios"))]
        if let tauri::RunEvent::Opened { urls } = _event {
            open_paths(_app, urls.into_iter().filter_map(|u| u.to_file_path().ok()));
        }
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decodes_encoded_paths() {
        assert_eq!(
            percent_decode("%2FUsers%2Fasha%2FBikri%20%E0%A4%B9%E0%A4%BF.xlsx").unwrap(),
            "/Users/asha/Bikri हि.xlsx"
        );
        assert_eq!(percent_decode("plain").unwrap(), "plain");
    }

    #[test]
    fn finds_files_in_arguments() {
        let args = vec!["app".into(), "--flag".into(), "a.xlsx".into(), "/x/b.csv".into()];
        let paths = paths_in_args(&args, Path::new("/home/u"));
        assert_eq!(paths, [PathBuf::from("/home/u/a.xlsx"), PathBuf::from("/x/b.csv")]);
    }
}
