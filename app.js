"use strict";

// Holds parsed data per file: { code -> summedQty }
const store = {
  bom: null,
  po: null,
  pr: null,
  onhand: null,
};

// Original filenames for status display
const fileNames = { bom: "", po: "", pr: "", onhand: "" };

const FILE_KEYS = ["bom", "po", "pr", "onhand"];

// Merged result rows, kept for download
let resultRows = [];

const els = {
  generate: document.getElementById("generateBtn"),
  downloadXlsx: document.getElementById("downloadXlsxBtn"),
  downloadCsv: document.getElementById("downloadCsvBtn"),
  reset: document.getElementById("resetBtn"),
  messages: document.getElementById("messages"),
  summary: document.getElementById("summary"),
  legend: document.getElementById("legend"),
  tbody: document.querySelector("#resultTable tbody"),
};

// ---- File input wiring ----------------------------------------------------

document.querySelectorAll("input[type=file][data-input]").forEach((input) => {
  const key = input.getAttribute("data-input");
  input.addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    handleFile(key, file);
  });
});

els.generate.addEventListener("click", generate);
els.downloadXlsx.addEventListener("click", () => downloadResult("xlsx"));
els.downloadCsv.addEventListener("click", () => downloadResult("csv"));
els.reset.addEventListener("click", resetAll);

// ---- Parsing --------------------------------------------------------------

function handleFile(key, file) {
  fileNames[key] = file.name;
  setStatus(key, "Reading…");
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const parsed = parseSheet(sheet);
      store[key] = parsed.map;
      setStatus(
        key,
        `${file.name} — ${parsed.uniqueCount} items` +
          (parsed.duplicates ? ` (${parsed.duplicates} dupes summed)` : ""),
        "loaded"
      );
      clearMessages();
      refreshButtons();
    } catch (err) {
      console.error(err);
      store[key] = null;
      setStatus(key, `Failed to read: ${err.message}`, "error");
      refreshButtons();
    }
  };
  reader.onerror = () => {
    store[key] = null;
    setStatus(key, "Failed to read file", "error");
    refreshButtons();
  };
  reader.readAsArrayBuffer(file);
}

// Reads column A (item code) and column B (qty). Row 1 is a header and skipped.
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
    const rawCode = row[0];
    const rawQty = row[1];
    const code = normalizeCode(rawCode);
    if (code === "") continue; // skip rows with no item code
    if (isTotalLabel(code)) continue; // skip pivot "Grand Total" / "Total" rows

    const qty = toNumber(rawQty);
    seenRows++;
    map.set(code, (map.get(code) || 0) + qty);
  }

  return {
    map,
    uniqueCount: map.size,
    duplicates: seenRows - map.size,
  };
}

function normalizeCode(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

// Pivot tables often append a summary row (e.g. "Grand Total").
// These are not real item codes, so skip them.
function isTotalLabel(code) {
  return /^(grand\s+total|total|sum|subtotal)$/i.test(code.trim());
}

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

// ---- Merge / generate -----------------------------------------------------

function generate() {
  // BOM is required, plus at least one of PO or PR.
  if (!store.bom) {
    showMessage("Engineering BOM is required to generate the table.", "err");
    return;
  }
  if (!store.po && !store.pr) {
    showMessage(
      "Provide at least one of PO stage or PR stage along with the BOM.",
      "err"
    );
    return;
  }

  const codes = new Set();
  [store.bom, store.po, store.pr, store.onhand].forEach((m) => {
    if (m) for (const code of m.keys()) codes.add(code);
  });

  const sorted = Array.from(codes).sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
  );

  const hasOnHand = !!store.onhand;

  resultRows = sorted.map((code) => {
    const bomQty = store.bom ? store.bom.get(code) || 0 : 0;
    const poQty = store.po ? store.po.get(code) || 0 : 0;
    const prQty = store.pr ? store.pr.get(code) || 0 : 0;
    const onHand = store.onhand ? store.onhand.get(code) || 0 : 0;
    const x = poQty + prQty;
    const y = bomQty - x;
    const status = computeStatus(y, onHand, hasOnHand);
    return { code, bomQty, poQty, prQty, x, y, onHand, status };
  });

  renderTable(resultRows, hasOnHand);
  renderSummary(resultRows, hasOnHand);
  els.legend.hidden = !hasOnHand;
  refreshButtons();

  const warnings = [];
  if (!store.po) warnings.push("PO stage not provided (PO Qty = 0)");
  if (!store.pr) warnings.push("PR stage not provided (PR Qty = 0)");
  if (!store.onhand) warnings.push("On Hand not provided (store status skipped)");
  if (warnings.length) showMessage(warnings.join(" · "), "warn");
}

// Store status only matters when there is a remaining requirement (Y > 0).
// Returns { level, shortfall } where level is red|yellow|green|null.
function computeStatus(y, onHand, hasOnHand) {
  if (!hasOnHand || y <= 0) return { level: null, shortfall: 0 };
  if (onHand <= 0) return { level: "red", shortfall: y };
  if (onHand < y) return { level: "yellow", shortfall: y - onHand };
  return { level: "green", shortfall: 0 };
}

function renderTable(rows, hasOnHand) {
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
  els.tbody.innerHTML = "";
  els.tbody.appendChild(frag);
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
      td.textContent = "—"; // Y <= 0, nothing to procure
  }
  return td;
}

function renderSummary(rows, hasOnHand) {
  const shortages = rows.filter((r) => r.y > 0).length;
  let text = `${rows.length} item(s). ${shortages} with remaining requirement (Y > 0).`;
  if (hasOnHand) {
    const red = rows.filter((r) => r.status.level === "red").length;
    const yellow = rows.filter((r) => r.status.level === "yellow").length;
    const green = rows.filter((r) => r.status.level === "green").length;
    text += ` Store: ${green} sufficient, ${yellow} partial, ${red} not in store.`;
  }
  els.summary.textContent = text;
}

function cell(text, mono) {
  const td = document.createElement("td");
  td.textContent = text;
  if (mono) td.style.fontFamily = "ui-monospace, monospace";
  return td;
}

function fmt(n) {
  if (typeof n !== "number" || isNaN(n)) return "0";
  // Keep integers clean, allow up to 4 decimals for fractional qtys.
  return Number.isInteger(n) ? String(n) : parseFloat(n.toFixed(4)).toString();
}

// ---- Download -------------------------------------------------------------

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

function statusText(status, hasOnHand) {
  if (!hasOnHand) return "";
  switch (status.level) {
    case "red": return "Not in store";
    case "yellow": return `Short by ${fmt(status.shortfall)}`;
    case "green": return "Sufficient";
    default: return "";
  }
}

function downloadResult(kind) {
  if (!resultRows.length) return;
  if (kind === "csv") downloadCsv();
  else downloadXlsx();
}

function stamp() {
  return new Date().toISOString().slice(0, 10);
}

function downloadCsv() {
  const hasOnHand = !!store.onhand;
  const lines = [HEADER.map(csvCell).join(",")];
  resultRows.forEach((r) => {
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
  triggerDownload(blob, `master_plan_${stamp()}.csv`);
}

function csvCell(v) {
  const s = v === null || v === undefined ? "" : String(v);
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

async function downloadXlsx() {
  const hasOnHand = !!store.onhand;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Master Plan");

  ws.addRow(HEADER);
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true };
  headerRow.alignment = { horizontal: "center" };

  resultRows.forEach((r) => {
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
      const statusCellRef = row.getCell(8);
      statusCellRef.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: FILL[r.status.level] },
      };
      statusCellRef.alignment = { horizontal: "center" };
      statusCellRef.font = { bold: true };
    }
  });

  // Reasonable column widths.
  ws.columns.forEach((col, i) => {
    col.width = i === 0 ? 18 : i === 7 ? 16 : 12;
  });
  ws.views = [{ state: "frozen", ySplit: 1 }];

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  triggerDownload(blob, `master_plan_${stamp()}.xlsx`);
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

// ---- UI helpers -----------------------------------------------------------

function setStatus(key, text, cls) {
  const el = document.querySelector(`[data-status="${key}"]`);
  if (!el) return;
  el.textContent = text;
  el.className = "file-status" + (cls ? " " + cls : "");
}

function refreshButtons() {
  // Need BOM plus at least one of PO or PR to generate.
  els.generate.disabled = !(store.bom && (store.po || store.pr));
  const hasResult = resultRows.length > 0;
  els.downloadXlsx.disabled = !hasResult;
  els.downloadCsv.disabled = !hasResult;
}

function showMessage(text, type) {
  const div = document.createElement("div");
  div.className = "msg " + (type || "info");
  div.textContent = text;
  els.messages.appendChild(div);
}

function clearMessages() {
  els.messages.innerHTML = "";
}

function resetAll() {
  store.bom = store.po = store.pr = store.onhand = null;
  resultRows = [];
  FILE_KEYS.forEach((k) => {
    fileNames[k] = "";
    setStatus(k, "No file");
    const input = document.querySelector(`input[data-input="${k}"]`);
    if (input) input.value = "";
  });
  els.tbody.innerHTML = "";
  els.summary.textContent = "";
  els.legend.hidden = true;
  clearMessages();
  refreshButtons();
}
