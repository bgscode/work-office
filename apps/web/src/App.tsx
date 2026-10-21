import { useEffect, useState } from "react";
import { loadEngine } from "@office/engine";
import { TitleBar, type AppKind } from "./chrome/TitleBar";
import { TitleSlot } from "./chrome/titleSlot";
import { SheetsApp } from "./sheets/SheetsApp";

export function App() {
  const [app, setApp] = useState<AppKind>("sheets");
  const [engine, setEngine] = useState<"loading" | "ready" | Error>("loading");
  const [titleSlot, setTitleSlot] = useState<HTMLElement | null>(null);

  useEffect(() => {
    loadEngine().then(
      () => setEngine("ready"),
      (err) => setEngine(err instanceof Error ? err : new Error(String(err))),
    );
  }, []);

  return (
    <div className="shell">
      <TitleBar app={app} onApp={setApp} slotRef={setTitleSlot} />
      <TitleSlot.Provider value={titleSlot}>
        {engine === "loading" && <p className="placeholder">Loading engine…</p>}
        {engine instanceof Error && <p className="placeholder">Engine failed to load: {engine.message}</p>}
        {engine === "ready" && app === "sheets" && <SheetsApp />}
        {engine === "ready" && app !== "sheets" && (
          <p className="placeholder">
            {app === "docs" ? "Docs" : "Slides"} is on the roadmap (docs/ROADMAP.md). Contributions welcome!
          </p>
        )}
      </TitleSlot.Provider>
    </div>
  );
}
