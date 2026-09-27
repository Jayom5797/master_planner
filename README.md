# Master Planner

A static, backend-free web app that merges an **Engineering BOM**, **PO stage**, and **PR stage** Excel/CSV into one procurement plan.

## What it does

Upload up to 3 files (each with **item code in column A** and **qty in column B**, header in row 1). The app produces a table:

| Column | Meaning |
| --- | --- |
| Item Code | Unique item code |
| BOM Qty | Quantity from the Engineering BOM |
| PO Qty | Quantity from the PO stage |
| PR Qty | Quantity from the PR stage |
| X (PO+PR) | PO Qty + PR Qty |
| Y (BOM−X) | BOM Qty − X (remaining requirement) |

Then download the result as **Excel (.xlsx)** or **CSV**.

### Rules
- Files are read by **column position** (A = code, B = qty), so exact header text doesn't matter.
- **Duplicate item codes** within a file are **summed**.
- An item missing from a file is treated as **qty 0**.
- Only the **BOM is required**; PO and PR are optional (missing → 0).
- Rows where `Y > 0` (still needs procurement) are highlighted.

## Tech
- Plain HTML/CSS/JS, no build step, no backend.
- [SheetJS](https://sheetjs.com/) via CDN for reading/writing spreadsheets.
- Runs entirely in the browser.

## Local use
Open `index.html` in a browser, or serve the folder with any static server.

## Deploy (Vercel)
This is a static site. In Vercel, import the repo and deploy with **no build command** and output directory = project root. `vercel.json` is included.
