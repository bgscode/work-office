import { useEffect, useState } from "react";
import { loadEngine } from "@office/engine";
import { SheetsApp } from "./sheets/SheetsApp";

type AppKind = "sheets" | "docs" | "slides";

export function App() {
  const [app, setApp] = useState<AppKind>("sheets");
  const [engine, setEngine] = useState<"loading" | "ready" | Error>("loading");

  useEffect(() => {
    loadEngine().then(
      () => setEngine("ready"),
      (err) => setEngine(err instanceof Error ? err : new Error(String(err))),
    );
  }, []);

  return (
    <div className="shell">
      <header className="app-bar">
        <span className="brand">Office</span>
        {(["sheets", "docs", "slides"] as const).map((k) => (
          <button key={k} className={k === app ? "active" : ""} onClick={() => setApp(k)}>
            {k}
          </button>
        ))}
      </header>
      {engine === "loading" && <p className="placeholder">Loading engine…</p>}
      {engine instanceof Error && <p className="placeholder">Engine failed to load: {engine.message}</p>}
      {engine === "ready" && app === "sheets" && <SheetsApp />}
      {engine === "ready" && app !== "sheets" && (
        <p className="placeholder">
          {app === "docs" ? "Docs" : "Slides"} is on the roadmap (docs/ROADMAP.md). Contributions welcome!
        </p>
      )}
    </div>
  );
}
