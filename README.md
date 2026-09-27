# Master Planner

A static, backend-free web app that merges an **Engineering BOM**, **PO stage**, and **PR stage** Excel/CSV into one procurement plan.

## What it does

Upload up to 4 files (each with **item code in column A** and **qty in column B**, header in row 1). The app produces a table:

| Column | Meaning |
| --- | --- |
| Item Code | Unique item code |
| BOM Qty | Quantity from the Engineering BOM |
| PO Qty | Quantity from the PO stage |
| PR Qty | Quantity from the PR stage |
| X (PO+PR) | PO Qty + PR Qty |
| Y (BOM−X) | BOM Qty − X (remaining requirement) |
| On Hand Qty | Quantity currently in store (from the On Hand file) |
| Store Status | Color-coded check of On Hand vs Y |

Then download the result as **Excel (.xlsx)** or **CSV**.

### Store Status (color-coded)
Only evaluated for items that still need procurement (`Y > 0`):

- 🔴 **Red** — On Hand = 0 (item not in store)
- 🟡 **Yellow** — On Hand < Y (partial; cell/label shows the shortfall)
- 🟢 **Green** — On Hand ≥ Y (sufficient in store)

Colors appear both in the on-screen table and as real cell fills in the downloaded Excel.

### Rules
- Files are read by **column position** (A = code, B = qty), so exact header text doesn't matter.
- **Duplicate item codes** within a file are **summed** (handles pivoted BOMs).
- Pivot summary rows like `Grand Total` / `Total` are ignored.
- An item missing from a file is treated as **qty 0**.
- **Generate requires the BOM plus at least one of PO or PR.** On Hand is optional.
- If On Hand is not provided, the Store Status column is left blank.

## Pages
- **Multi-File** (`index.html`) — one Excel/CSV per stage (BOM, PO, PR, On Hand).
- **Single File** (`single.html`) — one workbook with multiple sheets. You map each role (BOM/PO/PR/On Hand) to a sheet; it defaults to sheet order (Sheet1=BOM, Sheet2=PO, Sheet3=PR, Sheet4=On Hand) and lets you override.

Both pages share the same merge/status/download logic in `core.js` and are linked from a top navbar.

## Tech
- Plain HTML/CSS/JS, no build step, no backend.
- [SheetJS](https://sheetjs.com/) via CDN for reading spreadsheets.
- [ExcelJS](https://github.com/exceljs/exceljs) via CDN for writing color-filled Excel output.
- Runs entirely in the browser.

## Local use
Open `index.html` in a browser, or serve the folder with any static server.

## Deploy (Vercel)
This is a static site. In Vercel, import the repo and deploy with **no build command** and output directory = project root. `vercel.json` is included.
