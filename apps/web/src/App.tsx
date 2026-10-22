// Top level: Home page or an app, chosen by the URL hash (see router.ts).

import { useEffect, useState } from "react";
import { loadEngine } from "@office/engine";
import { TitleBar } from "./chrome/TitleBar";
import { TitleSlot } from "./chrome/titleSlot";
import { isDesktop, onFilesOpened, takeOpenedFiles } from "./desktop";
import { newDocumentId, type AppKind, type RecentDocument } from "./documents";
import { rememberOpenedFile } from "./files";
import { ComingSoon } from "./home/ComingSoon";
import { HomePage } from "./home/HomePage";
import { navigate, useRoute } from "./router";
import { SheetsLoader } from "./sheets/SheetsLoader";

/**
 * Files chosen on Home (or File > Open) waiting to be opened, by the new
 * document's id. Not React state: navigating renders the new route right away
 * (the hash change is handled as an urgent update), possibly before a state
 * update made just before it, and the loader must see the file on its first
 * render or it would open an empty workbook instead.
 */
const pendingFiles = new Map<string, File>();

export function App() {
  const route = useRoute();
  const [engine, setEngine] = useState<"loading" | "ready" | Error>("loading");
  const [titleSlot, setTitleSlot] = useState<HTMLElement | null>(null);

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
    pendingFiles.set(docId, file);
    rememberOpenedFile(docId, file);
    navigate(`/sheets/${docId}`);
  }

  // Desktop app: files double-clicked in Finder / Explorer, at launch and later.
  useEffect(() => {
    if (!isDesktop()) return;
    const openAll = () =>
      void takeOpenedFiles().then(
        (files) => files.forEach(openFile),
        (err) => console.error("Could not open the file", err),
      );
    openAll();
    return onFilesOpened(openAll);
  }, []);

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
          <SheetsLoader
            key={route.docId}
            docId={route.docId}
            initialFile={pendingFiles.get(route.docId)}
            onInitialFileOpened={() => pendingFiles.delete(route.docId!)}
            onNewWorkbook={() => newDocument("sheets")}
            onOpenFile={openFile}
          />
        )}
      </TitleSlot.Provider>
    </div>
  );
}
