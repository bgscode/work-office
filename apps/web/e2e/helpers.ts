// Shared helpers for the browser tests: open a fresh workbook, find cells on
// the canvas grid, use the ribbon, and fail on any page error.

import { expect, type Page } from "@playwright/test";

/** Excel's default column width and row height, and the header sizes. */
const COL = 64;
const ROW = 20;
const HEADER_W = 50;
const HEADER_H = 20;

export async function openWorkbook(page: Page) {
  const errors: string[] = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => m.type() === "error" && errors.push(m.text()));
  await page.goto("/#/sheets");
  await page.locator(".cell-editor").waitFor();
  return {
    errors,
    /** Screen point at the middle of a cell (0-based), for the default layout. */
    async cell(row: number, col: number) {
      const box = (await page.locator(".grid").boundingBox())!;
      return { x: box.x + HEADER_W + col * COL + COL / 2, y: box.y + HEADER_H + row * ROW + ROW / 2 };
    },
  };
}

export async function clickCell(page: Page, at: { x: number; y: number }) {
  await page.mouse.click(at.x, at.y);
}

/** Type rows of values starting at the active cell (Tab between, Enter after each row). */
export async function typeRows(page: Page, rows: string[][]) {
  for (const row of rows) {
    for (const v of row) {
      await page.keyboard.type(v);
      await page.keyboard.press("Tab");
    }
    await page.keyboard.press("Enter");
  }
}

export async function ribbonTab(page: Page, name: string) {
  await page.locator(".ribbon-tabs [role=tab]", { hasText: new RegExp(`^${name}$`) }).click();
}

export async function ribbonCommand(page: Page, label: string) {
  await page.locator(`.ribbon-commands button[aria-label="${label}"]`).first().click();
}

export async function menuItem(page: Page, label: string) {
  await page.locator(".ribbon-menu").getByText(label, { exact: true }).click();
}

export async function goTo(page: Page, ref: string) {
  await page.keyboard.press("Control+g");
  await page.keyboard.type(ref);
  await page.keyboard.press("Enter");
}

export const formulaBar = (page: Page) => page.locator("#formula-bar-input");

export async function expectNoErrors(errors: string[]) {
  expect(errors, errors.join("\n")).toEqual([]);
}
