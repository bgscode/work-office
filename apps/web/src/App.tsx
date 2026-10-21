// Top level: Home page or an app, chosen by the URL hash (see router.ts).

import { useEffect, useState } from "react";
import { loadEngine } from "@office/engine";
import { TitleBar } from "./chrome/TitleBar";
import { TitleSlot } from "./chrome/titleSlot";
import { newDocumentId, type AppKind, type RecentDocument } from "./documents";
import { ComingSoon } from "./home/ComingSoon";
import { HomePage } from "./home/HomePage";
import { navigate, useRoute } from "./router";
import { SheetsApp } from "./sheets/SheetsApp";

export function App() {
  const route = useRoute();
  const [engine, setEngine] = useState<"loading" | "ready" | Error>("loading");
  const [titleSlot, setTitleSlot] = useState<HTMLElement | null>(null);
  /** A file chosen on Home (or File > Open) waiting to be opened in a new document. */
  const [pendingFile, setPendingFile] = useState<{ docId: string; file: File } | null>(null);

  useEffect(() => {
    loadEngine().then(
      () => setEngine("ready"),
      (err) => setEngine(err instanceof Error ? err : new Error(String(err))),
    );
  }, []);

  // "#/sheets" with no document id means "new workbook": give it an id.
  useEffect(() => {
    if (route.page !== "home" && !route.docId) navigate(`/${route.page}/${newDocumentId()}`, true);
  }, [route]);

  const goHome = () => navigate("/");
  const newDocument = (app: AppKind) => navigate(`/${app}/${newDocumentId()}`);
  const openDocument = (doc: RecentDocument) => navigate(`/${doc.app}/${doc.id}`);
  function openFile(file: File) {
    const docId = newDocumentId();
    setPendingFile({ docId, file });
    navigate(`/sheets/${docId}`);
  }

  if (route.page === "home") {
    return <HomePage onOpenDocument={openDocument} onNew={newDocument} onOpenFile={openFile} />;
  }

  return (
    <div className="shell">
      <TitleBar app={route.page} onHome={goHome} slotRef={setTitleSlot} />
      <TitleSlot.Provider value={titleSlot}>
        {route.page !== "sheets" && <ComingSoon app={route.page} onHome={goHome} />}
        {route.page === "sheets" && engine === "loading" && <p className="placeholder">Loading…</p>}
        {route.page === "sheets" && engine instanceof Error && <p className="placeholder">Could not start: {engine.message}</p>}
        {route.page === "sheets" && engine === "ready" && route.docId && (
          <SheetsApp
            key={route.docId}
            docId={route.docId}
            initialFile={pendingFile?.docId === route.docId ? pendingFile.file : undefined}
            onInitialFileOpened={() => setPendingFile(null)}
            onNewWorkbook={() => newDocument("sheets")}
            onOpenFile={openFile}
          />
        )}
      </TitleSlot.Provider>
    </div>
  );
}
