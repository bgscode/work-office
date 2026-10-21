// Excel-style title bar: app mark, app switcher, a slot for the open
// document's name and save status (filled by the app through TitleSlot).

export type AppKind = "sheets" | "docs" | "slides";

const APPS: { kind: AppKind; label: string; letter: string; color: string }[] = [
  { kind: "sheets", label: "Sheets", letter: "X", color: "#107c41" },
  { kind: "docs", label: "Docs", letter: "W", color: "#185abd" },
  { kind: "slides", label: "Slides", letter: "P", color: "#c43e1c" },
];

interface Props {
  app: AppKind;
  onApp: (app: AppKind) => void;
  /** Receives the element the app portals its title into. */
  slotRef: (el: HTMLElement | null) => void;
}

export function TitleBar({ app, onApp, slotRef }: Props) {
  const current = APPS.find((a) => a.kind === app)!;
  return (
    <header className="app-bar title-bar" data-app={app}>
      <span className="app-mark" aria-hidden title={`Office ${current.label}`}>
        <span className="app-mark-tile" style={{ background: current.color }}>
          {current.letter}
        </span>
      </span>
      <nav className="app-switcher" aria-label="Apps">
        {APPS.map((a) => (
          <button key={a.kind} className={a.kind === app ? "active" : ""} aria-current={a.kind === app ? "page" : undefined} onClick={() => onApp(a.kind)}>
            {a.label}
          </button>
        ))}
      </nav>
      <div className="title-slot" ref={slotRef} />
      <span className="brand">Office</span>
    </header>
  );
}
