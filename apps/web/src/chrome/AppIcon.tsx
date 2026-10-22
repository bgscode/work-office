import type { AppKind } from "../documents";

export const APPS: Record<AppKind, { label: string; noun: string; color: string; available: boolean }> = {
  sheets: { label: "Sheets", noun: "workbook", color: "#217346", available: true },
  docs: { label: "Docs", noun: "document", color: "#2b579a", available: false },
  slides: { label: "Slides", noun: "presentation", color: "#c43e1c", available: false },
};

/** The app's tile: a coloured square with a simple white glyph (our own marks, not Microsoft's). */
export function AppIcon({ app, size = 24 }: { app: AppKind; size?: number }) {
  const glyph = {
    sheets: (
      <g stroke="#fff" strokeWidth="1.6" fill="none">
        <rect x="6" y="6" width="12" height="12" rx="1" />
        <path d="M6 10h12M6 14h12M10 6v12M14 6v12" strokeWidth="1.2" />
      </g>
    ),
    docs: (
      <g stroke="#fff" strokeWidth="1.8" strokeLinecap="round">
        <path d="M7 8h10M7 12h10M7 16h6" />
      </g>
    ),
    slides: (
      <g fill="none" stroke="#fff" strokeWidth="1.6">
        <rect x="6" y="7" width="12" height="9" rx="1" />
        <path d="M12 16v2.5M9.5 18.5h5" strokeLinecap="round" />
      </g>
    ),
  }[app];
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden="true" focusable="false">
      <rect width="24" height="24" rx="5" fill={APPS[app].color} />
      {glyph}
    </svg>
  );
}
