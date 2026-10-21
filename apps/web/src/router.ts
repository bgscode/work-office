// Tiny hash router: #/ (Home), #/sheets/<document id>, #/docs, #/slides.
// The hash keeps the browser's Back button and refresh working.

import { useEffect, useState } from "react";
import type { AppKind } from "./documents";

export type Route = { page: "home" } | { page: AppKind; docId?: string };

export function parseRoute(hash: string): Route {
  const [page, docId] = hash.replace(/^#\/?/, "").split("/");
  if (page === "sheets" || page === "docs" || page === "slides") return { page, docId: docId || undefined };
  return { page: "home" };
}

export function navigate(path: string, replace = false) {
  const url = `#${path}`;
  if (replace) history.replaceState(null, "", url);
  else history.pushState(null, "", url);
  window.dispatchEvent(new HashChangeEvent("hashchange"));
}

export function useRoute(): Route {
  const [route, setRoute] = useState(() => parseRoute(location.hash));
  useEffect(() => {
    const onChange = () => setRoute(parseRoute(location.hash));
    window.addEventListener("hashchange", onChange);
    window.addEventListener("popstate", onChange);
    return () => {
      window.removeEventListener("hashchange", onChange);
      window.removeEventListener("popstate", onChange);
    };
  }, []);
  return route;
}
