window.addEventListener("error", (e) => {
  const el = document.getElementById("msg");
  if (el) el.textContent = "JS Error: " + (e.message || e.error);
});

// ======================================================
// CHANGED: Replace GAS URL with your FastAPI local server
// OLD: const SCRIPT_URL = "https://script.google.com/macros/s/..."
// NEW: point to your FastAPI backend
// ======================================================
const API_URL = "http://127.0.0.1:8000";  // ← localhost while developing
                                           // ← swap to Render URL when deploying

let SHEET_BOUNDS = null;
let PLOT = null;

function extractSheetId(input) {
  const s = (input || "").trim();
  if (!s.startsWith("http")) return s;
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : s;
}

function setMsg(t) {
  document.getElementById("msg").textContent = t;
}

function setReadOut(t) {
  document.getElementById("readOut").textContent = t;
}

function setPlotInfo(t) {
  const el = document.getElementById("plotInfo");
  if (el) el.textContent = t;
}

// ======================================================
// CHANGED: Replace JSONP helper with simple fetch()
// OLD: used <script> tag injection (JSONP pattern)
// NEW: standard fetch() — works because FastAPI has CORS enabled
// ======================================================
async function apiFetch(endpoint, params) {
  const q = new URLSearchParams(params);
  const url = `${API_URL}/${endpoint}?${q.toString()}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP error: ${res.status}`);
  return res.json();
}

// Initialize dropdown
(function initDefaults() {
  const colSel = document.getElementById("colLetter");
  const plotSel = document.getElementById("plotCol");
  if (colSel && colSel.options.length === 0) colSel.add(new Option("A", "A"));
  if (plotSel && plotSel.options.length === 0) plotSel.add(new Option("A", "A"));
})();

async function ensureBounds(sheetId) {
  if (SHEET_BOUNDS && SHEET_BOUNDS.sheetId === sheetId) return SHEET_BOUNDS;

  // CHANGED: was jsonp({ action: "getBounds", sheetId })
  const res = await apiFetch("bounds", { sheet_id: sheetId });
  if (!res.ok) throw new Error(res.error || "getBounds failed");
  SHEET_BOUNDS = { sheetId, ...res };

  const rowInput = document.getElementById("rowNum");
  if (rowInput) rowInput.max = String(res.lastRow || 1);

  const A = "A".charCodeAt(0);
  const n = Math.max(1, Math.min(Number(res.lastCol || 1), 26));
  const headers = res.headers || [];

  function buildOptions() {
    const frag = document.createDocumentFragment();
    for (let i = 0; i < n; i++) {
      const col = String.fromCharCode(A + i);
      const h = headers[i] ? String(headers[i]) : "";
      frag.appendChild(new Option(h ? `${col} (${h})` : col, col));
    }
    return frag;
  }

  const colSel = document.getElementById("colLetter");
  if (colSel) {
    const prev = colSel.value || "A";
    colSel.replaceChildren(buildOptions());
    colSel.value = prev;
    if (!colSel.value) colSel.value = "A";
  }

  const plotSel = document.getElementById("plotCol");
  if (plotSel) {
    const prev = plotSel.value || "A";
    plotSel.replaceChildren(buildOptions());
    plotSel.value = prev;
    if (!plotSel.value) plotSel.value = "A";
  }

  return SHEET_BOUNDS;
}

// Plot type UI toggle — UNCHANGED
function updatePlotControls() {
  const t = document.getElementById("plotType")?.value || "bar";
  const bar = document.getElementById("barControls");
  const pie = document.getElementById("pieControls");
  if (bar) bar.style.display = (t === "bar") ? "inline" : "none";
  if (pie) pie.style.display = (t === "pie") ? "inline" : "none";
}

document.getElementById("plotType")?.addEventListener("change", updatePlotControls);
updatePlotControls();


// ===== Buttons =====

// Load bounds
document.getElementById("loadSheetBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  if (!sheetId) { setMsg("Please paste a sheet URL/ID first."); return; }
  try {
    setMsg("Loading sheet bounds...");
    const b = await ensureBounds(sheetId);
    const lastColLetter = String.fromCharCode("A".charCodeAt(0) + Math.min(b.lastCol || 1, 26) - 1);
    setMsg(`Bounds loaded: rows=1..${b.lastRow}, cols=A..${lastColLetter} (backend ms=${b.ms ?? "N/A"})`);
  } catch (e) {
    setMsg("Failed to load bounds: " + String(e));
  }
});

// (3) A+B -> C
document.getElementById("runBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }
  try {
    setMsg("Running A+B=C ...");
    // CHANGED: was jsonp({ action: "addCols", sheetId })
    const res = await apiFetch("add-cols", { sheet_id: sheetId });
    if (!res.ok) throw new Error(res.error || "unknown");
    setMsg(`${res.message || "Done."} (backend ms=${res.ms ?? "N/A"})`);
    SHEET_BOUNDS = null;
  } catch (e) {
    setMsg("Failed: " + String(e));
  }
});

// (4) Load cell
document.getElementById("loadBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  const row = parseInt(document.getElementById("rowNum").value, 10);
  const col = document.getElementById("colLetter").value;

  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }
  if (!row || row < 1) { setMsg("Row must be >= 1."); return; }
  if (!col) { setMsg("Please select a column."); return; }

  try {
    await ensureBounds(sheetId);
    setMsg("Loading cell...");
    // CHANGED: was jsonp({ action: "getCell", sheetId, row: String(row), col })
    const res = await apiFetch("cell", { sheet_id: sheetId, row: String(row), col });
    if (!res.ok) throw new Error(res.error || "unknown");

    setReadOut(
      `data point ${res.row} has ${res.featureName} (${res.value})\n` +
      `data type: ${res.type}\n` +
      `cell: ${res.col}${res.row}\n` +
      `backend ms: ${res.ms ?? "N/A"}`
    );
    setMsg("Loaded.");
  } catch (e) {
    setMsg("Load failed: " + String(e));
  }
});

// (4) Save cell
document.getElementById("saveBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  const row = parseInt(document.getElementById("rowNum").value, 10);
  const col = document.getElementById("colLetter").value;
  const value = document.getElementById("newValue").value;

  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }
  if (!row || row < 1) { setMsg("Row must be >= 1."); return; }
  if (!col) { setMsg("Please select a column."); return; }

  try {
    await ensureBounds(sheetId);
    setMsg("Saving...");
    // CHANGED: was jsonp({ action: "setCell", sheetId, row: String(row), col, value })
    const res = await apiFetch("set-cell", { sheet_id: sheetId, row: String(row), col, value });
    if (!res.ok) throw new Error(res.error || "unknown");

    setMsg(`Saved: ${col}${row} = ${res.writtenValue} (type=${res.writtenType}) (backend ms=${res.ms ?? "N/A"})`);
    SHEET_BOUNDS = null;

    // Re-read cell to confirm
    const reread = await apiFetch("cell", { sheet_id: sheetId, row: String(row), col });
    if (reread.ok) {
      setReadOut(
        `data point ${reread.row} has ${reread.featureName} (${reread.value})\n` +
        `data type: ${reread.type}\n` +
        `cell: ${reread.col}${reread.row}\n` +
        `backend ms: ${reread.ms ?? "N/A"}`
      );
    }
  } catch (e) {
    setMsg("Save failed: " + String(e));
  }
});

// ===== Plot helpers — ALL UNCHANGED (pure frontend logic) =====

function toNumberOrNull(x) {
  const s = String(x ?? "").trim();
  if (!s) return null;
  const v = Number(s);
  return Number.isFinite(v) ? v : null;
}

function isMostlyNumeric(arr) {
  let num = 0, non = 0;
  for (const v of arr) {
    const n = toNumberOrNull(v);
    if (n === null) non++; else num++;
  }
  return num >= non;
}

function buildCategoryCounts(arr, topK = 30) {
  const map = new Map();
  for (const v of arr) {
    const s = String(v ?? "").trim();
    if (!s) continue;
    map.set(s, (map.get(s) || 0) + 1);
  }
  const items = [...map.entries()].sort((a,b) => b[1]-a[1]).slice(0, topK);
  return { labels: items.map(x=>x[0]), counts: items.map(x=>x[1]) };
}

function buildNumericValueCounts(vals, topK = 30) {
  const map = new Map();
  for (const v of vals) {
    const n = toNumberOrNull(v);
    if (n === null) continue;
    const key = String(n);
    map.set(key, (map.get(key) || 0) + 1);
  }
  const items = [...map.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, Math.max(1, topK))
    .sort((a, b) => Number(a[0]) - Number(b[0]));
  return {
    labels: items.map(x => x[0]),
    counts: items.map(x => x[1]),
    uniqueCount: items.length
  };
}

function buildNumericDistribution(vals, N) {
  const vc = buildNumericValueCounts(vals, N);
  if (vc.uniqueCount > 0 && vc.uniqueCount <= N) {
    return { mode: "value", labels: vc.labels, counts: vc.counts, uniqueCount: vc.uniqueCount };
  }
  const h = buildHistogram(vals, N);
  return { mode: "hist", labels: h.labels, counts: h.counts, uniqueCount: vc.uniqueCount };
}

function buildHistogram(arr, bins = 8) {
  const nums = arr.map(toNumberOrNull).filter(v => v !== null);
  if (nums.length === 0) return { labels: [], counts: [] };
  let min = Math.min(...nums);
  let max = Math.max(...nums);
  if (min === max) { min -= 0.5; max += 0.5; }
  const width = (max - min) / bins;
  const counts = new Array(bins).fill(0);
  for (const x of nums) {
    let idx = Math.floor((x - min) / width);
    if (idx < 0) idx = 0;
    if (idx >= bins) idx = bins - 1;
    counts[idx]++;
  }
  const labels = [];
  for (let i = 0; i < bins; i++) {
    const a = min + i * width;
    const b = a + width;
    labels.push(`${a.toFixed(2)}–${b.toFixed(2)}`);
  }
  return { labels, counts };
}

function plotCountBar(labels, counts, title) {
  const canvas = document.getElementById("plotCanvas");
  const ctx = canvas.getContext("2d");
  if (PLOT) PLOT.destroy();
  PLOT = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets: [{ label: "count", data: counts }] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { title: { display: true, text: title } },
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
    }
  });
}

function buildPieCounts(arr, topK = 6, includeOthers = true) {
  const map = new Map();
  for (const v of arr) {
    const s = String(v ?? "").trim();
    if (!s) continue;
    map.set(s, (map.get(s) || 0) + 1);
  }
  const items = [...map.entries()].sort((a, b) => b[1] - a[1]);
  const top = items.slice(0, topK);
  const rest = items.slice(topK);
  const labels = top.map(x => x[0]);
  const counts = top.map(x => x[1]);
  if (includeOthers) {
    const othersCount = rest.reduce((sum, x) => sum + x[1], 0);
    if (othersCount > 0) { labels.push("Others"); counts.push(othersCount); }
  }
  return { labels, counts };
}

function plotPie(labels, counts, title) {
  const canvas = document.getElementById("plotCanvas");
  const ctx = canvas.getContext("2d");
  if (PLOT) PLOT.destroy();
  PLOT = new Chart(ctx, {
    type: "pie",
    data: { labels, datasets: [{ data: counts }] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: true, text: title },
        legend: { display: true, position: "bottom" }
      }
    }
  });
}

// Plot button
document.getElementById("plotBtn").addEventListener("click", async () => {
  const sheetId = extractSheetId(document.getElementById("sheetId").value);
  if (!sheetId) { setMsg("Please paste a sheet URL/ID."); return; }
  try {
    await ensureBounds(sheetId);
    const col = document.getElementById("plotCol")?.value || "A";
    const plotType = document.getElementById("plotType")?.value || "bar";
    setMsg(`Fetching column ${col}...`);
    setPlotInfo("");
    const t0 = performance.now();
    // CHANGED: was jsonp({ action: "getColumn", sheetId, col })
    const res = await apiFetch("column", { sheet_id: sheetId, col });
    const t1 = performance.now();
    if (!res.ok) throw new Error(res.error || "getColumn failed");
    const vals = res.values || [];
    const header = res.header || col;
    if (vals.length === 0) {
      setMsg("No data rows to plot.");
      if (PLOT) PLOT.destroy();
      return;
    }
    if (plotType === "bar") {
      const rawN = parseInt(document.getElementById("barNInput")?.value || "", 10);
      const N = Math.max(1, Math.min(Number.isFinite(rawN) ? rawN : 8, 60));
      if (isMostlyNumeric(vals)) {
        const h = buildHistogram(vals, N);
        plotCountBar(h.labels, h.counts, `${header} (histogram, bins=${N})`);
        setMsg(`Plotted BAR histogram for numeric column ${col}.`);
      } else {
        const c = buildCategoryCounts(vals, N);
        plotCountBar(c.labels, c.counts, `${header} (top ${N} categories)`);
        setMsg(`Plotted BAR category counts for column ${col}.`);
      }
    } else {
      const rawTopK = parseInt(document.getElementById("pieTopKInput")?.value || "", 10);
      const topK = Math.max(2, Math.min(Number.isFinite(rawTopK) ? rawTopK : 6, 20));
      const includeOthers = !!document.getElementById("pieOthersToggle")?.checked;
      if (isMostlyNumeric(vals)) {
        const dist = buildNumericDistribution(vals, topK);
        if (dist.mode === "value") {
          plotPie(dist.labels, dist.counts, `${header} (pie value-count, unique=${dist.uniqueCount})`);
          setMsg(`Plotted PIE value-count for numeric column ${col}.`);
        } else {
          plotPie(dist.labels, dist.counts, `${header} (pie histogram bins=${topK})`);
          setMsg(`Plotted PIE histogram for numeric column ${col}.`);
        }
      } else {
        const p = buildPieCounts(vals, topK, includeOthers);
        plotPie(p.labels, p.counts, `${header} (pie, top ${topK}${includeOthers ? " + Others" : ""})`);
        setMsg(`Plotted PIE for categorical column ${col}.`);
      }
    }
    setPlotInfo(`backend ms=${res.ms ?? "N/A"}, fetch+render ms=${Math.round(t1 - t0)}`);
  } catch (e) {
    setMsg("Plot failed: " + String(e));
  }
});

// Save PNG — UNCHANGED
document.getElementById("savePngBtn").addEventListener("click", () => {
  const canvas = document.getElementById("plotCanvas");
  const url = canvas.toDataURL("image/png");
  const a = document.createElement("a");
  a.href = url;
  a.download = "plot.png";
  a.click();
});

// Save PDF — UNCHANGED
document.getElementById("savePdfBtn").addEventListener("click", () => {
  const canvas = document.getElementById("plotCanvas");
  const imgData = canvas.toDataURL("image/png");
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const margin = 30;
  const imgW = pageW - margin * 2;
  const imgH = (canvas.height / canvas.width) * imgW;
  const y = (pageH - imgH) / 2;
  pdf.addImage(imgData, "PNG", margin, Math.max(margin, y), imgW, imgH);
  pdf.save("plot.pdf");
});
