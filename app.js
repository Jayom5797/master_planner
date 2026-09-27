"use strict";

// Multi-file page: one Excel/CSV per stage (BOM, PO, PR, On Hand).
// Shared logic lives in core.js (window.MasterPlanner).

const MP = window.MasterPlanner;

// Holds parsed data per file: { code -> summedQty }
const store = { bom: null, po: null, pr: null, onhand: null };
const FILE_KEYS = ["bom", "po", "pr", "onhand"];

let resultRows = [];
let resultHasOnHand = false;

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
els.downloadXlsx.addEventListener("click", () => doDownload("xlsx"));
els.downloadCsv.addEventListener("click", () => doDownload("csv"));
els.reset.addEventListener("click", resetAll);

// ---- Parsing --------------------------------------------------------------

async function handleFile(key, file) {
  setStatus(key, "Reading…");
  try {
    const wb = await MP.readWorkbook(file);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const parsed = MP.parseSheet(sheet);
    store[key] = parsed.map;
    setStatus(
      key,
      `${file.name} — ${parsed.uniqueCount} items` +
        (parsed.duplicates ? ` (${parsed.duplicates} dupes summed)` : ""),
      "loaded"
    );
    clearMessages();
  } catch (err) {
    console.error(err);
    store[key] = null;
    setStatus(key, `Failed to read: ${err.message}`, "error");
  }
  refreshButtons();
}

// ---- Generate -------------------------------------------------------------

function generate() {
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

  const { rows, hasOnHand } = MP.merge(store);
  resultRows = rows;
  resultHasOnHand = hasOnHand;

  MP.renderTable(els.tbody, rows, hasOnHand);
  els.summary.textContent = MP.summaryText(rows, hasOnHand);
  els.legend.hidden = !hasOnHand;
  refreshButtons();

  const warnings = [];
  if (!store.po) warnings.push("PO stage not provided (PO Qty = 0)");
  if (!store.pr) warnings.push("PR stage not provided (PR Qty = 0)");
  if (!store.onhand) warnings.push("On Hand not provided (store status skipped)");
  if (warnings.length) showMessage(warnings.join(" · "), "warn");
}

function doDownload(kind) {
  MP.download(kind, resultRows, resultHasOnHand, "master_plan");
}

// ---- UI helpers -----------------------------------------------------------

function setStatus(key, text, cls) {
  const el = document.querySelector(`[data-status="${key}"]`);
  if (!el) return;
  el.textContent = text;
  el.className = "file-status" + (cls ? " " + cls : "");
}

function refreshButtons() {
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
  resultHasOnHand = false;
  FILE_KEYS.forEach((k) => {
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
