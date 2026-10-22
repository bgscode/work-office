import { APPS, AppIcon } from "../chrome/AppIcon";
import type { AppKind } from "../documents";

/** `onHome` is left out in the Docs/Slides desktop apps, which have no Home yet. */
export function ComingSoon({ app, onHome }: { app: AppKind; onHome?: () => void }) {
  return (
    <div className="coming-soon">
      <AppIcon app={app} size={56} />
      <h1>{APPS[app].label} is coming soon</h1>
      <p>
        We're building {APPS[app].label} next, after Sheets. This app updates itself, so it will turn into {APPS[app].label} as soon as it's
        ready. Follow progress in <code>docs/ROADMAP.md</code> — contributions are welcome.
      </p>
      {onHome && <button onClick={onHome}>Back to Home</button>}
    </div>
  );
}
