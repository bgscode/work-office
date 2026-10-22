import { APPS, AppIcon } from "../chrome/AppIcon";
import type { AppKind } from "../documents";

export function ComingSoon({ app, onHome }: { app: AppKind; onHome: () => void }) {
  return (
    <div className="coming-soon">
      <AppIcon app={app} size={56} />
      <h1>{APPS[app].label} is coming soon</h1>
      <p>
        We're building {APPS[app].label} next, after Sheets. Follow progress in <code>docs/ROADMAP.md</code> — contributions are welcome.
      </p>
      <button onClick={onHome}>Back to Home</button>
    </div>
  );
}
