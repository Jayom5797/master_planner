"use strict";

/*
 * MasterPlanner core — shared, page-agnostic logic used by every page.
 * Pure functions for parsing sheets, merging BOM/PO/PR/On-Hand, computing
 * store status, plus reusable rendering and download helpers.
 *
 * Exposed as a global (window.MasterPlanner) so plain <script> pages can use
 * it with no build step.
 */
(function (global) {
  const HEADER = [
    "Item Code",
    "BOM Qty",
    "PO Qty",
    "PR Qty",
    "X (PO+PR)",
    "Y (BOM-X)",
    "On Hand Qty",
    "Store Status",
  ];

  // ARGB fills for the Store Status cell in Excel.
  const FILL = {
    red: "FFEF4444",
    yellow: "FFF59E0B",
    green: "FF22C55E",
  };

  // ---- Value helpers ------------------------------------------------------

  function normalizeCode(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  // Pivot tables often append a summary row (e.g. "Grand Total").
  function isTotalLabel(code) {
    return /^(grand\s+total|total|sum|subtotal)$/i.test(String(code).trim());
  }

  function toNumber(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    const n = parseFloat(String(v).replace(/,/g, "").trim());
    return isNaN(n) ? 0 : n;
  }

  function fmt(n) {
    if (typeof n !== "number" || isNaN(n)) return "0";
    return Number.isInteger(n) ? String(n) : parseFloat(n.toFixed(4)).toString();
  }

  // ---- BOM pivot helpers --------------------------------------------------

  // A row is dropped if the ENGG REMARKS says the item was deleted. This must
  // tolerate any case, extra words ("deleted by team", "party to delete"), and
  // common misspellings of "delete" (deletd, delted, deleated, deleteed, ...).
  // Strategy: strip to letters, then look for a d-e?-l-e?-(a)?-t stem, which
  // matches delete/deleted/deletd/delted/deleated etc. without matching safe
  // remarks like "added", "model is updated", "r01".
  function isDeletedRemark(remark) {
    const s = String(remark == null ? "" : remark).toLowerCase();
    if (!s.trim()) return false;
    // Normalise: keep only a-z and spaces.
    const letters = s.replace(/[^a-z ]/g, " ");
    return /\bde+l+e*a*t/.test(letters) || /\bdle+t/.test(letters);
  }

  // Piping BOMs are identified by "pip" appearing in the sheet name (covers
  // both "PIPING" and the misspelled "PIPNG").
  function isPipingSheetName(name) {
    return /pip/i.test(String(name));
  }

  // Find the header row of a BOM sheet: the first row (within the first 20)
  // that contains both an "ERP CODE" and a "QTY" cell. Returns -1 if none.
  function findBomHeaderRow(rows) {
    const limit = Math.min(rows.length, 20);
    for (let i = 0; i < limit; i++) {
      const r = (rows[i] || []).map((c) => String(c).trim().toUpperCase());
      if (r.includes("ERP CODE") && r.includes("QTY")) return i;
    }
    return -1;
  }

  // Parse one BOM sheet: locate columns by header name (ERP CODE, QTY,
  // ENGG REMARKS), then accumulate qty per code into `map`, skipping blank
  // codes, total labels, and deleted rows.
  // Returns { ok, kept, deleted } — ok=false when no header row was found.
  function accumulateBomSheet(sheet, map) {
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: "",
      blankrows: false,
    });
    const h = findBomHeaderRow(rows);
    if (h < 0) return { ok: false, kept: 0, deleted: 0 };

    const hdr = (rows[h] || []).map((c) => String(c).trim());
    const erpIdx = hdr.findIndex((c) => c.toUpperCase() === "ERP CODE");
    const qtyIdx = hdr.findIndex((c) => c.toUpperCase() === "QTY");
    const remIdx = hdr.findIndex((c) => c.toUpperCase().includes("ENGG REMARK"));

    let kept = 0;
    let deleted = 0;
    for (let i = h + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const code = normalizeCode(row[erpIdx]);
      if (code === "" || isTotalLabel(code)) continue;
      if (remIdx >= 0 && isDeletedRemark(row[remIdx])) {
        deleted++;
        continue;
      }
      map.set(code, (map.get(code) || 0) + toNumber(row[qtyIdx]));
      kept++;
    }
    return { ok: true, kept, deleted };
  }

  // Classify every sheet in a workbook and build two pivot maps.
  // Returns {
  //   piping: Map<code,qty>, nonPiping: Map<code,qty>,
  //   sheets: [{ name, group: 'piping'|'nonPiping'|'skipped', kept, deleted }]
  // }
  function pivotBoms(workbook) {
    const piping = new Map();
    const nonPiping = new Map();
    const sheets = [];

    workbook.SheetNames.forEach((name) => {
      const ws = workbook.Sheets[name];
      const target = isPipingSheetName(name) ? piping : nonPiping;
      const res = accumulateBomSheet(ws, target);
      if (!res.ok) {
        sheets.push({ name, group: "skipped", kept: 0, deleted: 0 });
      } else {
        sheets.push({
          name,
          group: isPipingSheetName(name) ? "piping" : "nonPiping",
          kept: res.kept,
          deleted: res.deleted,
        });
      }
    });

    return { piping, nonPiping, sheets };
  }

  // Turn a Map<code,qty> into a sorted [{ code, qty }] array.
  function mapToSortedRows(map) {
    return Array.from(map.entries())
      .sort((a, b) =>
        a[0].localeCompare(b[0], undefined, { numeric: true, sensitivity: "base" })
      )
      .map(([code, qty]) => ({ code, qty }));
  }

  // ---- Parsing ------------------------------------------------------------

  // Reads column A (item code) and column B (qty). Row 1 is a header, skipped.
  // Duplicate codes are summed. Returns { map, uniqueCount, duplicates }.
  function parseSheet(sheet) {
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: "",
      blankrows: false,
    });

    const map = new Map();
    let seenRows = 0;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const code = normalizeCode(row[0]);
      if (code === "") continue;
      if (isTotalLabel(code)) continue;

      seenRows++;
      map.set(code, (map.get(code) || 0) + toNumber(row[1]));
    }

    return { map, uniqueCount: map.size, duplicates: seenRows - map.size };
  }

  // Read a File (Blob) into an XLSX workbook. Returns a Promise<workbook>.
  function readWorkbook(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const data = new Uint8Array(ev.target.result);
          resolve(XLSX.read(data, { type: "array" }));
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error("Failed to read file"));
      reader.readAsArrayBuffer(file);
    });
  }

  // ---- Merge / status -----------------------------------------------------

  // Store status only matters when there is a remaining requirement (Y > 0).
  // Returns { level, shortfall } where level is red|yellow|green|null.
  function computeStatus(y, onHand, hasOnHand) {
    if (!hasOnHand || y <= 0) return { level: null, shortfall: 0 };
    if (onHand <= 0) return { level: "red", shortfall: y };
    if (onHand < y) return { level: "yellow", shortfall: y - onHand };
    return { level: "green", shortfall: 0 };
  }

  // maps: { bom, po, pr, onhand } — each a Map<code, qty> or null.
  // Rows come from BOM/PO/PR only; On Hand is a lookup source, not a row source.
  // Returns { rows, hasOnHand }.
  function merge(maps) {
    const { bom, po, pr, onhand } = maps;
    const hasOnHand = !!onhand;

    const codes = new Set();
    [bom, po, pr].forEach((m) => {
      if (m) for (const code of m.keys()) codes.add(code);
    });

    const sorted = Array.from(codes).sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
    );

    const rows = sorted.map((code) => {
      const bomQty = bom ? bom.get(code) || 0 : 0;
      const poQty = po ? po.get(code) || 0 : 0;
      const prQty = pr ? pr.get(code) || 0 : 0;
      const onHand = onhand ? onhand.get(code) || 0 : 0;
      const x = poQty + prQty;
      const y = bomQty - x;
      const status = computeStatus(y, onHand, hasOnHand);
      return { code, bomQty, poQty, prQty, x, y, onHand, status };
    });

    return { rows, hasOnHand };
  }

  function statusText(status, hasOnHand) {
    if (!hasOnHand) return "";
    switch (status.level) {
      case "red": return "Not in store";
      case "yellow": return `Short by ${fmt(status.shortfall)}`;
      case "green": return "Sufficient";
      default: return "";
    }
  }

  // ---- Rendering (DOM) ----------------------------------------------------

  function cell(text, mono) {
    const td = document.createElement("td");
    td.textContent = text;
    if (mono) td.style.fontFamily = "ui-monospace, monospace";
    return td;
  }

  function statusCell(status, hasOnHand) {
    const td = document.createElement("td");
    td.className = "status";
    if (!hasOnHand) {
      td.classList.add("none");
      td.textContent = "—";
      return td;
    }
    switch (status.level) {
      case "red":
        td.classList.add("red");
        td.textContent = "Not in store";
        break;
      case "yellow":
        td.classList.add("yellow");
        td.textContent = `Short by ${fmt(status.shortfall)}`;
        break;
      case "green":
        td.classList.add("green");
        td.textContent = "Sufficient";
        break;
      default:
        td.classList.add("none");
        td.textContent = "—";
    }
    return td;
  }

  function renderTable(tbody, rows, hasOnHand) {
    const frag = document.createDocumentFragment();
    rows.forEach((r) => {
      const tr = document.createElement("tr");
      tr.appendChild(cell(r.code, true));
      tr.appendChild(cell(fmt(r.bomQty)));
      tr.appendChild(cell(fmt(r.poQty)));
      tr.appendChild(cell(fmt(r.prQty)));
      tr.appendChild(cell(fmt(r.x)));
      tr.appendChild(cell(fmt(r.y)));
      tr.appendChild(cell(hasOnHand ? fmt(r.onHand) : "—"));
      tr.appendChild(statusCell(r.status, hasOnHand));
      frag.appendChild(tr);
    });
    tbody.innerHTML = "";
    tbody.appendChild(frag);
  }

  function summaryText(rows, hasOnHand) {
    const shortages = rows.filter((r) => r.y > 0).length;
    let text = `${rows.length} item(s). ${shortages} with remaining requirement (Y > 0).`;
    if (hasOnHand) {
      const red = rows.filter((r) => r.status.level === "red").length;
      const yellow = rows.filter((r) => r.status.level === "yellow").length;
      const green = rows.filter((r) => r.status.level === "green").length;
      text += ` Store: ${green} sufficient, ${yellow} partial, ${red} not in store.`;
    }
    return text;
  }

  // ---- Download -----------------------------------------------------------

  function stamp() {
    return new Date().toISOString().slice(0, 10);
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function csvCell(v) {
    const s = v === null || v === undefined ? "" : String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function downloadCsv(rows, hasOnHand, filenameBase) {
    const lines = [HEADER.map(csvCell).join(",")];
    rows.forEach((r) => {
      lines.push(
        [
          r.code,
          r.bomQty,
          r.poQty,
          r.prQty,
          r.x,
          r.y,
          hasOnHand ? r.onHand : "",
          statusText(r.status, hasOnHand),
        ]
          .map(csvCell)
          .join(",")
      );
    });
    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], {
      type: "text/csv;charset=utf-8;",
    });
    triggerDownload(blob, `${filenameBase}_${stamp()}.csv`);
  }

  async function downloadXlsx(rows, hasOnHand, filenameBase) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Master Plan");

    ws.addRow(HEADER);
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: "center" };

    rows.forEach((r) => {
      const row = ws.addRow([
        r.code,
        r.bomQty,
        r.poQty,
        r.prQty,
        r.x,
        r.y,
        hasOnHand ? r.onHand : null,
        statusText(r.status, hasOnHand),
      ]);

      if (hasOnHand && r.status.level) {
        const sc = row.getCell(8);
        sc.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: FILL[r.status.level] },
        };
        sc.alignment = { horizontal: "center" };
        sc.font = { bold: true };
      }
    });

    ws.columns.forEach((col, i) => {
      col.width = i === 0 ? 18 : i === 7 ? 16 : 12;
    });
    ws.views = [{ state: "frozen", ySplit: 1 }];

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    triggerDownload(blob, `${filenameBase}_${stamp()}.xlsx`);
  }

  function download(kind, rows, hasOnHand, filenameBase) {
    if (!rows || !rows.length) return;
    const base = filenameBase || "master_plan";
    if (kind === "csv") downloadCsv(rows, hasOnHand, base);
    else downloadXlsx(rows, hasOnHand, base);
  }

  // ---- Pivot downloads ----------------------------------------------------

  const PIVOT_HEADER = ["Item Code", "Qty"];

  // Download both pivot tables in one workbook (two sheets).
  async function downloadPivotXlsx(nonPipingRows, pipingRows, filenameBase) {
    const wb = new ExcelJS.Workbook();

    const addSheet = (title, rows) => {
      const ws = wb.addWorksheet(title);
      ws.addRow(PIVOT_HEADER);
      ws.getRow(1).font = { bold: true };
      rows.forEach((r) => ws.addRow([r.code, r.qty]));
      ws.columns = [{ width: 20 }, { width: 12 }];
      ws.views = [{ state: "frozen", ySplit: 1 }];
    };

    addSheet("Non-Piping BOM", nonPipingRows);
    addSheet("Piping BOM", pipingRows);

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    triggerDownload(blob, `${filenameBase || "master_bom"}_${stamp()}.xlsx`);
  }

  // Download a single pivot table as CSV.
  function downloadPivotCsv(rows, filenameBase) {
    const lines = [PIVOT_HEADER.map(csvCell).join(",")];
    rows.forEach((r) => lines.push([r.code, r.qty].map(csvCell).join(",")));
    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], {
      type: "text/csv;charset=utf-8;",
    });
    triggerDownload(blob, `${filenameBase || "master_bom"}_${stamp()}.csv`);
  }

  // ---- Public API ---------------------------------------------------------

  global.MasterPlanner = {
    HEADER,
    parseSheet,
    readWorkbook,
    computeStatus,
    merge,
    statusText,
    fmt,
    renderTable,
    summaryText,
    download,
    triggerDownload,
    // BOM pivot
    isDeletedRemark,
    isPipingSheetName,
    pivotBoms,
    mapToSortedRows,
    downloadPivotXlsx,
    downloadPivotCsv,
  };
})(window);
