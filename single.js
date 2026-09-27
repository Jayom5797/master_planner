"use strict";

// Single-file page: one workbook, multiple sheets. The user maps each role
// (BOM / PO / PR / On Hand) to a sheet, then we run the shared merge logic.

const MP = window.MasterPlanner;
const ROLES = ["bom", "po", "pr", "onhand"];

let workbook = null; // XLSX workbook
let sheetNames = [];
let resultRows = [];
let resultHasOnHand = false;

const els = {
  input: document.getElementById("workbookInput"),
  status: document.getElementById("workbookStatus"),
  mapping: document.getElementById("mapping"),
  generate: document.getElementById("generateBtn"),
  downloadXlsx: document.getElementById("downloadXlsxBtn"),
  downloadCsv: document.getElementById("downloadCsvBtn"),
  reset: document.getElementById("resetBtn"),
  messages: document.getElementById("messages"),
  summary: document.getElementById("summary"),
  legend: document.getElementById("legend"),
  tbody: document.querySelector("#resultTable tbody"),
  downloadTemplate: document.getElementById("downloadTemplateBtn"),
};

const selects = {};
ROLES.forEach((role) => {
  selects[role] = document.querySelector(`select[data-role="${role}"]`);
  selects[role].addEventListener("change", refreshButtons);
});

els.input.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (file) loadWorkbook(file);
});
els.generate.addEventListener("click", generate);
els.downloadXlsx.addEventListener("click", () => doDownload("xlsx"));
els.downloadCsv.addEventListener("click", () => doDownload("csv"));
els.reset.addEventListener("click", resetAll);
els.downloadTemplate.addEventListener("click", downloadTemplate);

// ---- Blank template -------------------------------------------------------

// Empty workbook with four sheets (BOM, PO, PR, ON HAND) and no headers —
// the user pastes columns straight from their ERP, headers included.
async function downloadTemplate() {
  const wb = new ExcelJS.Workbook();
  ["BOM", "PO", "PR", "ON HAND"].forEach((name) => wb.addWorksheet(name));
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  MP.triggerDownload(blob, "master_planner_template.xlsx");
}

// ---- Load workbook & build the sheet mapping ------------------------------

async function loadWorkbook(file) {
  setStatus("Reading…");
  try {
    workbook = await MP.readWorkbook(file);
    sheetNames = workbook.SheetNames.slice();
    if (!sheetNames.length) throw new Error("No sheets found");

    buildSelects();
    applyDefaultMapping();
    els.mapping.hidden = false;
    setStatus(`${file.name} — ${sheetNames.length} sheet(s)`, "loaded");
    clearMessages();
  } catch (err) {
    console.error(err);
    workbook = null;
    sheetNames = [];
    els.mapping.hidden = true;
    setStatus(`Failed to read: ${err.message}`, "error");
  }
  refreshButtons();
}

function buildSelects() {
  ROLES.forEach((role) => {
    const sel = selects[role];
    sel.innerHTML = "";
    const none = document.createElement("option");
    none.value = "";
    none.textContent = "None";
    sel.appendChild(none);
    sheetNames.forEach((name, i) => {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = `${i + 1}. ${name}`;
      sel.appendChild(opt);
    });
  });
}

// Default: sheet order = BOM, PO, PR, On Hand (only if that many sheets exist).
function applyDefaultMapping() {
  const order = ["bom", "po", "pr", "onhand"];
  order.forEach((role, i) => {
    selects[role].value = i < sheetNames.length ? String(i) : "";
  });
}

// ---- Generate -------------------------------------------------------------

function getMappedSheetMap(role) {
  const val = selects[role].value;
  if (val === "") return null;
  const idx = parseInt(val, 10);
  const name = sheetNames[idx];
  const sheet = workbook.Sheets[name];
  if (!sheet) return null;
  return MP.parseSheet(sheet).map;
}

function generate() {
  if (!workbook) {
    showMessage("Load a workbook first.", "err");
    return;
  }

  // Guard against mapping two roles to the same sheet — usually a mistake.
  const chosen = ROLES.map((r) => selects[r].value).filter((v) => v !== "");
  const dupSheet = chosen.length !== new Set(chosen).size;

  const maps = {
    bom: getMappedSheetMap("bom"),
    po: getMappedSheetMap("po"),
    pr: getMappedSheetMap("pr"),
    onhand: getMappedSheetMap("onhand"),
  };

  if (!maps.bom) {
    showMessage("Map a sheet to the Engineering BOM role.", "err");
    return;
  }
  if (!maps.po && !maps.pr) {
    showMessage("Map at least one of PO or PR (plus BOM).", "err");
    return;
  }

  const { rows, hasOnHand } = MP.merge(maps);
  resultRows = rows;
  resultHasOnHand = hasOnHand;

  MP.renderTable(els.tbody, rows, hasOnHand);
  els.summary.textContent = MP.summaryText(rows, hasOnHand);
  els.legend.hidden = !hasOnHand;
  refreshButtons();

  const warnings = [];
  if (dupSheet) warnings.push("Same sheet mapped to more than one role");
  if (!maps.po) warnings.push("PO not mapped (PO Qty = 0)");
  if (!maps.pr) warnings.push("PR not mapped (PR Qty = 0)");
  if (!maps.onhand) warnings.push("On Hand not mapped (store status skipped)");
  if (warnings.length) showMessage(warnings.join(" · "), "warn");
}

function doDownload(kind) {
  MP.download(kind, resultRows, resultHasOnHand, "master_plan");
}

// ---- UI helpers -----------------------------------------------------------

function setStatus(text, cls) {
  els.status.textContent = text;
  els.status.className = "file-status" + (cls ? " " + cls : "");
}

function refreshButtons() {
  const bomSet = workbook && selects.bom.value !== "";
  const poOrPr =
    workbook && (selects.po.value !== "" || selects.pr.value !== "");
  els.generate.disabled = !(bomSet && poOrPr);
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
  workbook = null;
  sheetNames = [];
  resultRows = [];
  resultHasOnHand = false;
  els.input.value = "";
  els.mapping.hidden = true;
  ROLES.forEach((r) => (selects[r].innerHTML = ""));
  setStatus("No file");
  els.tbody.innerHTML = "";
  els.summary.textContent = "";
  els.legend.hidden = true;
  clearMessages();
  refreshButtons();
}
