// Last line of defence: if rendering throws, or the WebAssembly engine
// crashes (a Rust panic, see `installPanicHook` in crates/wasm), show a calm
// recovery screen instead of a blank page or a frozen grid. Every edit is
// already journaled to IndexedDB (sheets/journal.ts), so reloading picks up
// where the user was.

import { Component, type ErrorInfo, type ReactNode } from "react";
import { isEngineCrash, onEnginePanic } from "@office/engine";

interface State {
  error: Error | null;
}

export class ErrorBoundary extends Component<{ children: ReactNode }, State> {
  state: State = { error: null };
  private unsubscribe: (() => void) | null = null;

  static getDerivedStateFromError(error: unknown): State {
    return { error: error instanceof Error ? error : new Error(String(error)) };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("Office crashed while rendering:", error, info.componentStack);
  }

  componentDidMount() {
    this.unsubscribe = onEnginePanic((message) => this.crash(new Error(message)));
    window.addEventListener("error", this.onWindowError);
    window.addEventListener("unhandledrejection", this.onRejection);
  }

  componentWillUnmount() {
    this.unsubscribe?.();
    window.removeEventListener("error", this.onWindowError);
    window.removeEventListener("unhandledrejection", this.onRejection);
  }

  // Errors thrown in event handlers don't reach error boundaries; an engine
  // crash there still has to stop the app, since the next call would fail too.
  private onWindowError = (e: ErrorEvent) => {
    if (isEngineCrash(e.error ?? e.message)) this.crash(e.error instanceof Error ? e.error : new Error(e.message));
  };

  private onRejection = (e: PromiseRejectionEvent) => {
    if (isEngineCrash(e.reason)) this.crash(e.reason instanceof Error ? e.reason : new Error(String(e.reason)));
  };

  private crash(error: Error) {
    if (!this.state.error) this.setState({ error });
  }

  render() {
    const { error } = this.state;
    if (!error) return this.props.children;
    return <CrashScreen error={error} />;
  }
}

export function CrashScreen({ error }: { error: Error }) {
  return (
    <div className="crash-screen" role="alertdialog" aria-modal="true" aria-labelledby="crash-title" aria-describedby="crash-text">
      <div className="crash-card">
        <h1 id="crash-title">Something went wrong</h1>
        <p id="crash-text">
          Your work is saved in this browser as you go. Reload to pick up where you left off; at most the last change may be
          missing.
        </p>
        <div className="crash-actions">
          <button className="dialog-button primary" autoFocus onClick={() => window.location.reload()}>
            Reload
          </button>
          <button
            className="dialog-button"
            onClick={() => {
              window.location.hash = "#/";
              window.location.reload();
            }}
          >
            Go to Home
          </button>
        </div>
        <details>
          <summary>Technical details</summary>
          <pre>{error.stack ?? error.message}</pre>
        </details>
      </div>
    </div>
  );
}
