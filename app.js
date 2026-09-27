"use strict";

// Holds parsed data per file: { code -> summedQty }
const store = {
  bom: null,
  po: null,
  pr: null,
};

// Original filenames for status display
const fileNames = { bom: "", po: "", pr: "" };

// Merged result rows, kept for download
let resultRows = [];

const els = {
  generate: document.getElementById("generateBtn"),
  downloadXlsx: document.getElementById("downloadXlsxBtn"),
  downloadCsv: document.getElementById("downloadCsvBtn"),
  reset: document.getElementById("resetBtn"),
  messages: document.getElementById("messages"),
  summary: document.getElementById("summary"),
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

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

// ---- Merge / generate -----------------------------------------------------

function generate() {
  if (!store.bom) {
    showMessage("Engineering BOM is required to generate the table.", "err");
    return;
  }

  const codes = new Set();
  [store.bom, store.po, store.pr].forEach((m) => {
    if (m) for (const code of m.keys()) codes.add(code);
  });

  const sorted = Array.from(codes).sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
  );

  resultRows = sorted.map((code) => {
    const bomQty = store.bom ? store.bom.get(code) || 0 : 0;
    const poQty = store.po ? store.po.get(code) || 0 : 0;
    const prQty = store.pr ? store.pr.get(code) || 0 : 0;
    const x = poQty + prQty;
    const y = bomQty - x;
    return { code, bomQty, poQty, prQty, x, y };
  });

  renderTable(resultRows);
  renderSummary(resultRows);
  refreshButtons();

  const warnings = [];
  if (!store.po) warnings.push("PO stage not provided (PO Qty = 0)");
  if (!store.pr) warnings.push("PR stage not provided (PR Qty = 0)");
  if (warnings.length) showMessage(warnings.join(" · "), "warn");
}

function renderTable(rows) {
  const frag = document.createDocumentFragment();
  rows.forEach((r) => {
    const tr = document.createElement("tr");
    if (r.y > 0) tr.classList.add("shortage"); // still needed to procure
    tr.appendChild(cell(r.code, true));
    tr.appendChild(cell(fmt(r.bomQty)));
    tr.appendChild(cell(fmt(r.poQty)));
    tr.appendChild(cell(fmt(r.prQty)));
    tr.appendChild(cell(fmt(r.x)));
    tr.appendChild(cell(fmt(r.y)));
    frag.appendChild(tr);
  });
  els.tbody.innerHTML = "";
  els.tbody.appendChild(frag);
}

function renderSummary(rows) {
  const shortages = rows.filter((r) => r.y > 0).length;
  els.summary.textContent =
    `${rows.length} item(s). ` +
    `${shortages} with remaining requirement (Y > 0).`;
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

function downloadResult(kind) {
  if (!resultRows.length) return;

  const header = [
    "Item Code",
    "BOM Qty",
    "PO Qty",
    "PR Qty",
    "X (PO+PR)",
    "Y (BOM-X)",
  ];
  const aoa = [header].concat(
    resultRows.map((r) => [r.code, r.bomQty, r.poQty, r.prQty, r.x, r.y])
  );

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Master Plan");

  const stamp = new Date().toISOString().slice(0, 10);
  const name = `master_plan_${stamp}.${kind}`;
  XLSX.writeFile(wb, name, { bookType: kind });
}

// ---- UI helpers -----------------------------------------------------------

function setStatus(key, text, cls) {
  const el = document.querySelector(`[data-status="${key}"]`);
  if (!el) return;
  el.textContent = text;
  el.className = "file-status" + (cls ? " " + cls : "");
}

function refreshButtons() {
  els.generate.disabled = !store.bom;
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
  store.bom = store.po = store.pr = null;
  resultRows = [];
  ["bom", "po", "pr"].forEach((k) => {
    fileNames[k] = "";
    setStatus(k, "No file");
    const input = document.querySelector(`input[data-input="${k}"]`);
    if (input) input.value = "";
  });
  els.tbody.innerHTML = "";
  els.summary.textContent = "";
  clearMessages();
  refreshButtons();
}
