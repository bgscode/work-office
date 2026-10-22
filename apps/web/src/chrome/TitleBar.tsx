// The app's title bar, like Excel's: a button back to Home, the app icon, and
// a slot where the open app shows its quick-access buttons, file name and save
// status (filled by the app through TitleSlot).

import type { AppKind } from "../documents";
import { APPS, AppIcon } from "./AppIcon";

interface Props {
  app: AppKind;
  onHome: () => void;
  /** Receives the element the app portals its title into. */
  slotRef: (el: HTMLElement | null) => void;
}

export function TitleBar({ app, onHome, slotRef }: Props) {
  return (
    <header className="title-bar" data-tauri-drag-region data-app={app} style={{ ["--app-color" as string]: APPS[app].color }}>
      <button className="home-button" onClick={onHome} aria-label="Go to Home" title="Home">
        <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
          {[2, 7, 12].flatMap((x) => [2, 7, 12].map((y) => <rect key={`${x}${y}`} x={x - 1} y={y - 1} width="3" height="3" rx="0.5" fill="#fff" />))}
        </svg>
      </button>
      <span className="title-app" data-tauri-drag-region>
        <AppIcon app={app} size={20} />
        <span>{APPS[app].label}</span>
      </span>
      <div className="title-slot" data-tauri-drag-region ref={slotRef} />
    </header>
  );
}
