// The start page, like the Microsoft 365 home: an app rail on the left, a
// greeting, "Start new" tiles, and the recent documents saved in this browser.

import { useEffect, useMemo, useState, type DragEvent } from "react";
import { APPS, AppIcon } from "../chrome/AppIcon";
import { listRecent, removeDocument, type AppKind, type RecentDocument } from "../documents";
import { pickFile } from "../files";

interface Props {
  onOpenDocument: (doc: RecentDocument) => void;
  onNew: (app: AppKind) => void;
  /** Open a .xlsx/.csv file from the computer. */
  onOpenFile: (file: File) => void;
}

const APP_ORDER: AppKind[] = ["sheets", "docs", "slides"];

function greeting(now: Date): string {
  const h = now.getHours();
  return h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening";
}

/** "Just now", "5 minutes ago", "Yesterday", or a date. */
export function relativeTime(then: number, now: number): string {
  const minutes = Math.round((now - then) / 60_000);
  if (minutes < 1) return "Just now";
  if (minutes < 60) return `${minutes} minute${minutes === 1 ? "" : "s"} ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours} hour${hours === 1 ? "" : "s"} ago`;
  const days = Math.round(hours / 24);
  if (days === 1) return "Yesterday";
  if (days < 7) return `${days} days ago`;
  return new Date(then).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
}

export function HomePage({ onOpenDocument, onNew, onOpenFile }: Props) {
  const [recent, setRecent] = useState(listRecent);
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<AppKind | "all">("all");
  const [dragging, setDragging] = useState(false);
  const now = Date.now();

  useEffect(() => {
    document.title = "Office";
    // Re-read after mount: a workbook being closed saves in its cleanup,
    // which runs after this page's first render.
    setRecent(listRecent());
  }, []);

  const shown = useMemo(
    () =>
      recent.filter(
        (d) => (filter === "all" || d.app === filter) && d.name.toLowerCase().includes(query.trim().toLowerCase()),
      ),
    [recent, query, filter],
  );

  async function open() {
    const file = await pickFile(".xlsx,.csv");
    if (file) onOpenFile(file);
  }

  function onDrop(e: DragEvent) {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && /\.(xlsx|csv)$/i.test(file.name)) onOpenFile(file);
  }

  return (
    <div
      className={`home${dragging ? " dragging" : ""}`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={(e) => e.currentTarget === e.target && setDragging(false)}
      onDrop={onDrop}
    >
      <header className="home-top">
        <span className="home-brand">Office</span>
        <label className="home-search">
          <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
            <circle cx="7" cy="7" r="4.5" fill="none" stroke="currentColor" strokeWidth="1.4" />
            <path d="M10.5 10.5 14 14" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
          </svg>
          <input type="search" placeholder="Search your files" aria-label="Search your files" value={query} onChange={(e) => setQuery(e.target.value)} />
        </label>
        <span className="home-top-spacer" />
      </header>

      <div className="home-body">
        <nav className="home-rail" aria-label="Apps">
          <button className="active" aria-current="page" title="Home">
            <svg width="22" height="22" viewBox="0 0 24 24" aria-hidden="true">
              <path d="M4 11 12 4l8 7v9h-5v-6H9v6H4z" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" />
            </svg>
            <span>Home</span>
          </button>
          <button onClick={open} title="Open a file">
            <svg width="22" height="22" viewBox="0 0 24 24" aria-hidden="true">
              <path d="M3 7h6l2 2h10v10H3z" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" />
            </svg>
            <span>Open</span>
          </button>
          <hr />
          {APP_ORDER.map((app) => (
            <button key={app} onClick={() => onNew(app)} title={APPS[app].available ? `New ${APPS[app].noun}` : `${APPS[app].label} — coming soon`}>
              <AppIcon app={app} size={28} />
              <span>{APPS[app].label}</span>
            </button>
          ))}
        </nav>

        <main className="home-main">
          <h1>{greeting(new Date(now))}</h1>

          <section aria-labelledby="start-new">
            <h2 id="start-new">Start new</h2>
            <div className="home-tiles">
              {APP_ORDER.map((app) => (
                <button key={app} className="home-tile" onClick={() => onNew(app)} disabled={!APPS[app].available}>
                  <span className={`tile-preview ${app}`} aria-hidden>
                    {app === "sheets" && <span className="tile-grid" />}
                    {app === "docs" && <span className="tile-lines" />}
                    {app === "slides" && <span className="tile-slide" />}
                  </span>
                  <span className="tile-label">
                    <AppIcon app={app} size={16} />
                    Blank {APPS[app].noun}
                    {!APPS[app].available && <em>Coming soon</em>}
                  </span>
                </button>
              ))}
              <button className="home-tile open-tile" onClick={open}>
                <span className="tile-preview open" aria-hidden>
                  <svg width="40" height="40" viewBox="0 0 24 24">
                    <path d="M3 7h6l2 2h10v10H3z" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
                    <path d="M12 17v-6m-2.5 2.5L12 11l2.5 2.5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
                  </svg>
                </span>
                <span className="tile-label">Open .xlsx or .csv</span>
              </button>
            </div>
          </section>

          <section aria-labelledby="recent">
            <div className="home-section-head">
              <h2 id="recent">Recent</h2>
              <div className="home-filters" role="tablist" aria-label="Filter by app">
                {(["all", ...APP_ORDER] as const).map((f) => (
                  <button key={f} role="tab" aria-selected={filter === f} className={filter === f ? "active" : ""} onClick={() => setFilter(f)}>
                    {f === "all" ? "All" : APPS[f].label}
                  </button>
                ))}
              </div>
            </div>
            {shown.length === 0 ? (
              <p className="home-empty">
                {recent.length === 0
                  ? "Nothing here yet. Create a blank workbook, or open or drop an .xlsx/.csv file anywhere on this page."
                  : "No files match."}
              </p>
            ) : (
              <table className="home-recent">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Saved</th>
                    <th aria-label="Actions" />
                  </tr>
                </thead>
                <tbody>
                  {shown.map((d) => (
                    <tr key={d.id} onClick={() => onOpenDocument(d)} tabIndex={0} onKeyDown={(e) => e.key === "Enter" && onOpenDocument(d)}>
                      <td>
                        <span className="recent-name">
                          <AppIcon app={d.app} size={20} />
                          {d.name}
                        </span>
                      </td>
                      <td className="recent-time">{relativeTime(d.savedAt, now)}</td>
                      <td className="recent-actions">
                        <button
                          aria-label={`Remove ${d.name} from this browser`}
                          title="Remove from this browser"
                          onClick={(e) => {
                            e.stopPropagation();
                            if (!confirm(`Remove "${d.name}" from this browser? This can't be undone.`)) return;
                            void removeDocument(d.id).then(() => setRecent(listRecent()));
                          }}
                        >
                          ×
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
            <p className="home-note">Files are saved in this browser automatically. Download a copy from File to keep it elsewhere.</p>
          </section>
        </main>
      </div>
      {dragging && <div className="home-drop">Drop the file to open it</div>}
    </div>
  );
}
