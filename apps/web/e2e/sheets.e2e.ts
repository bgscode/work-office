// Main Sheets flows, end to end: typing and formulas, formatting and undo,
// data tools, charts, shapes, freeze panes, print, saving across reloads,
// damaged files and the crash screen.

import { expect, test } from "@playwright/test";
import { clickCell, expectNoErrors, formulaBar, goTo, menuItem, openWorkbook, ribbonCommand, ribbonTab, typeRows } from "./helpers";

const TABLE = [
  ["Month", "Sales", "Cost"],
  ["Jan", "120", "80"],
  ["Feb", "150", "90"],
  ["Mar", "90", "70"],
  ["Apr", "200", "110"],
];

test("typing, formulas and autocomplete", async ({ page }) => {
  const { errors, cell } = await openWorkbook(page);
  await clickCell(page, await cell(0, 0));
  await typeRows(page, [["10"], ["20"], ["30"]]);
  await page.keyboard.type("=su");
  await expect(page.getByRole("option", { name: /^SUM/ }).first()).toBeVisible();
  await page.keyboard.press("Tab");
  await page.keyboard.type("A1:A3");
  await page.keyboard.press("Enter");
  await clickCell(page, await cell(3, 0));
  await expect(formulaBar(page)).toHaveValue("=SUM(A1:A3)");
  // Changing a precedent recalculates.
  await clickCell(page, await cell(0, 0));
  await page.keyboard.type("100");
  await page.keyboard.press("Enter");
  await clickCell(page, await cell(0, 0));
  await page.keyboard.down("Shift");
  await clickCell(page, await cell(3, 0));
  await page.keyboard.up("Shift");
  await expect(page.locator(".status-bar")).toContainText("Sum: 300");
  await expectNoErrors(errors);
});

test("formatting and undo", async ({ page }) => {
  const { errors, cell } = await openWorkbook(page);
  await clickCell(page, await cell(0, 0));
  await page.keyboard.type("Title");
  await page.keyboard.press("Enter");
  await clickCell(page, await cell(0, 0));
  await page.keyboard.press("Control+b");
  await expect(page.locator('.ribbon-commands button[aria-label="Bold"]')).toHaveAttribute("aria-pressed", "true");
  await page.keyboard.press("Control+z");
  await expect(page.locator('.ribbon-commands button[aria-label="Bold"]')).toHaveAttribute("aria-pressed", "false");
  await expectNoErrors(errors);
});

test("filter, sort and validation", async ({ page }) => {
  const { errors, cell } = await openWorkbook(page);
  await clickCell(page, await cell(0, 0));
  await typeRows(page, TABLE);
  await clickCell(page, await cell(1, 1));
  await page.keyboard.press("Control+Shift+L");
  // The filter button sits at the bottom-right of the header cell.
  const b1 = await cell(0, 1);
  await page.mouse.click(b1.x + 23, b1.y + 1);
  await page.locator(".filter-menu-check", { hasText: "90" }).click();
  await page.locator(".filter-menu").getByRole("button", { name: "OK" }).click();
  await expect(page.locator(".status-bar")).toContainText("3 of 4 records found");

  await ribbonTab(page, "Data");
  await ribbonCommand(page, "Clear");
  await clickCell(page, await cell(1, 3));
  await ribbonCommand(page, "Data Validation");
  await page.getByRole("combobox", { name: "Allow" }).click();
  await page.getByRole("option", { name: "List" }).click();
  await page.getByLabel("Source").fill("Yes,No");
  await page.getByRole("button", { name: "OK", exact: true }).click();
  await clickCell(page, await cell(1, 3));
  await page.keyboard.type("Maybe");
  await page.keyboard.press("Enter");
  await expect(page.locator(".validation-alert")).toBeVisible();
  await page.locator(".validation-alert").getByRole("button", { name: "Cancel" }).click();
  await expectNoErrors(errors);
});

test("charts and shapes", async ({ page }) => {
  const { errors, cell } = await openWorkbook(page);
  await clickCell(page, await cell(0, 0));
  await typeRows(page, TABLE);
  await clickCell(page, await cell(1, 1));
  await ribbonTab(page, "Insert");
  await ribbonCommand(page, "Column");
  await expect(page.locator(".drawing")).toHaveCount(1);
  await expect(page.locator(".ribbon-tabs [aria-selected=true]")).toHaveText("Chart Design");
  await ribbonCommand(page, "Change Chart Type");
  await menuItem(page, "Pie");
  await page.keyboard.press("Delete");
  await expect(page.locator(".drawing")).toHaveCount(0);
  await page.keyboard.press("Control+z");
  await expect(page.locator(".drawing")).toHaveCount(1);

  await ribbonTab(page, "Insert");
  await ribbonCommand(page, "Shapes");
  await page.getByRole("button", { name: "Oval", exact: true }).click();
  const from = await cell(12, 1);
  const to = await cell(18, 4);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator(".drawing")).toHaveCount(2);
  await page.keyboard.type("Hi");
  await page.keyboard.press("Escape");
  await expect(page.locator(".ribbon-tabs [aria-selected=true]")).toHaveText("Shape Format");
  await expectNoErrors(errors);
});

test("freeze panes while scrolled keeps the grid usable", async ({ page }) => {
  const { errors } = await openWorkbook(page);
  await goTo(page, "Z60");
  await ribbonTab(page, "View");
  await ribbonCommand(page, "Freeze Panes");
  await menuItem(page, "Freeze Panes");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowRight");
  await expect(page.getByLabel("Cell AA61")).toBeAttached();
  await ribbonCommand(page, "Freeze Panes");
  await menuItem(page, "Unfreeze Panes");
  await expectNoErrors(errors);
});

test("print preview and page layout", async ({ page }) => {
  const { errors, cell } = await openWorkbook(page);
  await clickCell(page, await cell(0, 0));
  await typeRows(page, TABLE);
  await ribbonTab(page, "Page Layout");
  await ribbonCommand(page, "Orientation");
  await menuItem(page, "Landscape");
  await page.keyboard.press("Control+p");
  await expect(page.locator(".print-nav")).toContainText("1 of 1");
  await expect(page.getByRole("combobox", { name: "Orientation" })).toContainText("Landscape");
  await page.keyboard.press("Escape");
  await expect(page.locator(".print-view")).toHaveCount(0);
  await expectNoErrors(errors);
});

test("work survives a reload", async ({ page }) => {
  const { errors, cell } = await openWorkbook(page);
  await clickCell(page, await cell(0, 0));
  await page.keyboard.type("kept");
  await page.keyboard.press("Enter");
  await expect(page.locator(".save-status")).toHaveText("Saved");
  await page.reload();
  await page.locator(".cell-editor").waitFor();
  await clickCell(page, await cell(0, 0));
  await expect(formulaBar(page)).toHaveValue("kept");
  await expectNoErrors(errors);
});

test("a damaged file shows a clear message", async ({ page }) => {
  await page.goto("/#/");
  const chooser = page.waitForEvent("filechooser");
  await page.getByText("Open .xlsx or .csv").click();
  await (await chooser).setFiles({ name: "broken.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from("not a workbook") });
  await expect(page.getByText(/isn't a valid \.xlsx workbook/)).toBeVisible();
});

test("an engine crash shows the recovery screen", async ({ page }) => {
  await openWorkbook(page);
  await page.evaluate(() => {
    try {
      (window as unknown as { __officeTest: { panic: () => void } }).__officeTest.panic();
    } catch {
      // The panic surfaces as an error; the boundary handles it.
    }
  });
  await expect(page.getByRole("alertdialog")).toContainText("Something went wrong");
  await expect(page.getByRole("button", { name: "Reload" })).toBeVisible();
});

test("chart title in place, pie colours and a narrow ribbon", async ({ page }) => {
  const { errors, cell } = await openWorkbook(page);
  await clickCell(page, await cell(0, 0));
  await typeRows(page, TABLE.map((r) => r.slice(0, 2)));
  await clickCell(page, await cell(1, 1));
  await ribbonTab(page, "Insert");
  await ribbonCommand(page, "Pie");
  const chart = page.locator(".drawing").first();
  const box = (await chart.boundingBox())!;
  await page.mouse.dblclick(box.x + box.width / 2, box.y + 20);
  await page.getByLabel("Chart title").fill("Sales by month");
  await page.keyboard.press("Enter");
  await ribbonCommand(page, "Change Colors");
  await menuItem(page, "Monochromatic Green");
  await page.keyboard.press("Control+z");
  await page.keyboard.press("Control+y");

  // Narrow window: groups fold into buttons that open their commands.
  await page.setViewportSize({ width: 1024, height: 700 });
  await ribbonTab(page, "Home");
  const alignment = page.locator(".ribbon-group.folded", { hasText: "Alignment" });
  await expect(alignment).toBeVisible();
  await alignment.getByRole("button", { name: "Alignment" }).click();
  await expect(page.locator(".folded-popup").getByRole("button", { name: "Wrap Text" })).toBeVisible();
  await expectNoErrors(errors);
});
