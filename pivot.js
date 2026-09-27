"use strict";

// BOM Pivot page: one workbook with many BOM sheets -> two master pivots
// (Non-Piping and Piping), each summed by ERP CODE, ignoring deleted rows.

const MP = window.MasterPlanner;

let workbook = null;
let nonPipingRows = [];
let pipingRows = [];

const els = {
  input: document.getElementById("workbookInput"),
  status: document.getElementById("workbookStatus"),
  generate: document.getElementById("generateBtn"),
  downloadXlsx: document.getElementById("downloadXlsxBtn"),
  reset: document.getElementById("resetBtn"),
  messages: document.getElementById("messages"),
  sheetReport: document.getElementById("sheetReport"),
  sheetTbody: document.querySelector("#sheetTable tbody"),
  nonPipingTbody: document.querySelector("#nonPipingTable tbody"),
  pipingTbody: document.querySelector("#pipingTable tbody"),
  nonPipingSummary: document.getElementById("nonPipingSummary"),
  pipingSummary: document.getElementById("pipingSummary"),
  csvNonPiping: document.getElementById("csvNonPipingBtn"),
  csvPiping: document.getElementById("csvPipingBtn"),
};

els.input.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (file) loadWorkbook(file);
});
els.generate.addEventListener("click", generate);
els.downloadXlsx.addEventListener("click", () =>
  MP.downloadPivotXlsx(nonPipingRows, pipingRows, "master_bom")
);
els.csvNonPiping.addEventListener("click", () =>
  MP.downloadPivotCsv(nonPipingRows, "master_bom_non_piping")
);
els.csvPiping.addEventListener("click", () =>
  MP.downloadPivotCsv(pipingRows, "master_bom_piping")
);
els.reset.addEventListener("click", resetAll);

// ---- Load -----------------------------------------------------------------

async function loadWorkbook(file) {
  setStatus("Reading…");
  try {
    workbook = await MP.readWorkbook(file);
    if (!workbook.SheetNames.length) throw new Error("No sheets found");
    setStatus(`${file.name} — ${workbook.SheetNames.length} sheet(s)`, "loaded");
    clearMessages();
  } catch (err) {
    console.error(err);
    workbook = null;
    setStatus(`Failed to read: ${err.message}`, "error");
  }
  refreshButtons();
}

// ---- Generate -------------------------------------------------------------

function generate() {
  if (!workbook) {
    showMessage("Load a workbook first.", "err");
    return;
  }

  const { piping, nonPiping, sheets } = MP.pivotBoms(workbook);
  nonPipingRows = MP.mapToSortedRows(nonPiping);
  pipingRows = MP.mapToSortedRows(piping);

  renderSheetReport(sheets);
  renderPivot(els.nonPipingTbody, nonPipingRows);
  renderPivot(els.pipingTbody, pipingRows);

  const nDeleted = sheets.reduce((s, x) => s + x.deleted, 0);
  els.nonPipingSummary.textContent = `${nonPipingRows.length} unique item(s).`;
  els.pipingSummary.textContent = `${pipingRows.length} unique item(s).`;

  refreshButtons();

  const skipped = sheets.filter((s) => s.group === "skipped").map((s) => s.name);
  const notes = [];
  if (skipped.length) notes.push(`Skipped (no ERP CODE/QTY header): ${skipped.join(", ")}`);
  if (nDeleted) notes.push(`${nDeleted} deleted row(s) ignored`);
  if (!pipingRows.length) notes.push("No piping sheets found");
  if (!nonPipingRows.length) notes.push("No non-piping sheets found");
  if (notes.length) showMessage(notes.join(" · "), "warn");
}

function renderPivot(tbody, rows) {
  const frag = document.createDocumentFragment();
  rows.forEach((r) => {
    const tr = document.createElement("tr");
    const c1 = document.createElement("td");
    c1.textContent = r.code;
    c1.style.fontFamily = "ui-monospace, monospace";
    const c2 = document.createElement("td");
    c2.textContent = MP.fmt(r.qty);
    tr.appendChild(c1);
    tr.appendChild(c2);
    frag.appendChild(tr);
  });
  tbody.innerHTML = "";
  tbody.appendChild(frag);
}

function renderSheetReport(sheets) {
  const frag = document.createDocumentFragment();
  sheets.forEach((s) => {
    const tr = document.createElement("tr");
    const label =
      s.group === "piping" ? "Piping" :
      s.group === "nonPiping" ? "Non-Piping" : "Skipped";
    [s.name, label, s.group === "skipped" ? "—" : String(s.kept),
     s.group === "skipped" ? "—" : String(s.deleted)].forEach((v, i) => {
      const td = document.createElement("td");
      td.textContent = v;
      if (i === 0) td.style.textAlign = "left";
      tr.appendChild(td);
    });
    if (s.group === "skipped") tr.style.opacity = "0.55";
    frag.appendChild(tr);
  });
  els.sheetTbody.innerHTML = "";
  els.sheetTbody.appendChild(frag);
  els.sheetReport.hidden = false;
}

// ---- UI helpers -----------------------------------------------------------

function setStatus(text, cls) {
  els.status.textContent = text;
  els.status.className = "file-status" + (cls ? " " + cls : "");
}

function refreshButtons() {
  els.generate.disabled = !workbook;
  els.downloadXlsx.disabled = !(nonPipingRows.length || pipingRows.length);
  els.csvNonPiping.disabled = !nonPipingRows.length;
  els.csvPiping.disabled = !pipingRows.length;
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
  nonPipingRows = [];
  pipingRows = [];
  els.input.value = "";
  setStatus("No file");
  els.sheetReport.hidden = true;
  els.sheetTbody.innerHTML = "";
  els.nonPipingTbody.innerHTML = "";
  els.pipingTbody.innerHTML = "";
  els.nonPipingSummary.textContent = "";
  els.pipingSummary.textContent = "";
  clearMessages();
  refreshButtons();
}
