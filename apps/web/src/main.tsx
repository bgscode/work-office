import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { triggerEnginePanicForTesting } from "@office/engine";
import { App } from "./App";
import { ErrorBoundary } from "./ErrorBoundary";
import { markMacDesktop } from "./desktop";
import "./styles.css";
import "./excel.css";
import "./home/home.css";

// Test hooks for the browser tests (apps/web/e2e); not in production builds.
if (import.meta.env.DEV) {
  (window as unknown as { __officeTest: object }).__officeTest = { panic: triggerEnginePanicForTesting };
}

markMacDesktop();

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
);
