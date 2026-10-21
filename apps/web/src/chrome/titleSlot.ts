// The title bar is drawn by App, but its middle (workbook name, save status,
// quick-access buttons) belongs to the open app. App provides an element here
// and the app renders into it with a portal (see sheets/WorkbookTitle.tsx).

import { createContext } from "react";

export const TitleSlot = createContext<HTMLElement | null>(null);
