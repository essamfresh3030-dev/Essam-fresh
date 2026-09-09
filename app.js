const fileInput = document.getElementById("fileInput");
const dropzone = document.getElementById("dropzone");
const emptyState = document.getElementById("emptyState");
const dashboard = document.getElementById("dashboard");
const sheetSelect = document.getElementById("sheetSelect");
const catCol = document.getElementById("catCol");
const valCol = document.getElementById("valCol");
const dateCol = document.getElementById("dateCol");
const search = document.getElementById("search");
const headerRowSelect = document.getElementById("headerRowSelect");

let workbook = null;
let rawGrid = [];
let rows = [];
let headers = [];
let charts = { bar: null, pie: null, line: null };

dropzone.addEventListener("click", () => fileInput.click());
["dragover", "dragenter"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.add("drag");
  });
});
["dragleave", "drop"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("drag");
  });
});
dropzone.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files[0];
  if (f) loadFile(f);
});
fileInput.addEventListener("change", (e) => {
  const f = e.target.files[0];
  if (f) loadFile(f);
});

function loadFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    workbook = XLSX.read(e.target.result, { type: "array", cellDates: true });
    document.getElementById("fileName").textContent = file.name;
    sheetSelect.innerHTML = workbook.SheetNames.map(
      (n) => `<option value="${n}">${n}</option>`
    ).join("");
    emptyState.classList.add("hidden");
    dashboard.classList.remove("hidden");
    useSheet(workbook.SheetNames[0]);
  };
  reader.readAsArrayBuffer(file);
}

sheetSelect.addEventListener("change", () => useSheet(sheetSelect.value, true));
headerRowSelect.addEventListener("change", applyHeaderRow);
[catCol, valCol, dateCol, search].forEach((el) =>
  el.addEventListener("input", renderAll)
);
document.getElementById("onlyInvalid").addEventListener("change", renderAll);
document.getElementById("onlySelectedCols").addEventListener("change", renderAll);
document.getElementById("applyRule").addEventListener("click", applyCurrentRule);
document.getElementById("clearRules").addEventListener("click", () => {
  colRules = {};
  renderRulesList();
  renderAll();
});

function escapeHtml(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function cellPreview(v) {
  if (v instanceof Date) return v.toLocaleDateString("ar-EG");
  const s = String(v ?? "").replace(/\s+/g, " ").trim();
  return s.length > 18 ? s.slice(0, 18) + "…" : s;
}

function guessHeaderRow(grid) {
  let best = 0;
  let bestScore = -1;
  const limit = Math.min(grid.length, 30);
  for (let i = 0; i < limit; i++) {
    const row = grid[i] || [];
    const texts = row.filter((v) => v !== "" && v != null && typeof v !== "number" && !(v instanceof Date));
    const nums = row.filter((v) => typeof v === "number");
    const filled = row.filter((v) => v !== "" && v != null).length;
    const score = texts.length * 3 + filled - nums.length * 2;
    if (score > bestScore) {
      bestScore = score;
      best = i;
    }
  }
  return best;
}

function fillHeaderRowSelect(preferred) {
  const limit = Math.min(rawGrid.length, 50);
  headerRowSelect.innerHTML = Array.from({ length: limit }, (_, i) => {
    const preview = (rawGrid[i] || []).slice(0, 4).map(cellPreview).filter(Boolean).join(" | ");
    return `<option value="${i}">صف ${i + 1}${preview ? " — " + preview : ""}</option>`;
  }).join("");
  headerRowSelect.value = String(Math.min(preferred, Math.max(limit - 1, 0)));
}

function applyHeaderRow() {
  const idx = Number(headerRowSelect.value) || 0;
  const maxLen = rawGrid.reduce((m, line) => Math.max(m, (line || []).length), 0);
  const headerCells = [...(rawGrid[idx] || [])];
  while (headerCells.length < maxLen) headerCells.push("");
  const used = new Map();
  headers = headerCells.map((h, i) => {
    let name = String(h ?? "").trim() || `عمود ${i + 1}`;
    const n = (used.get(name) || 0) + 1;
    used.set(name, n);
    return n > 1 ? `${name} (${n})` : name;
  });
  rows = rawGrid.slice(idx + 1).map((line) => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = line && line[i] !== undefined ? line[i] : "";
    });
    return obj;
  });
  document.getElementById("rowCount").textContent = `${rows.length} صف`;
  document.getElementById("colCount").textContent = `${headers.length} عمود`;
  fillSelects();
  fillRuleCol();
  renderAll();
}

function useSheet(name, resetHeader = true) {
  const sheet = workbook.Sheets[name];
  rawGrid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true, blankrows: true });
  if (!rawGrid.length) rawGrid = [[]];
  const preferred = resetHeader ? guessHeaderRow(rawGrid) : Number(headerRowSelect.value) || 0;
  fillHeaderRowSelect(preferred);
  applyHeaderRow();
}

function isNumericCol(h) {
  const sample = rows.slice(0, 40).map((r) => r[h]).filter((v) => v !== "" && v != null);
  if (!sample.length) return false;
  const n = sample.filter((v) => typeof v === "number" || /^-?\d+(\.\d+)?$/.test(String(v))).length;
  return n / sample.length > 0.7;
}

function isDateCol(h) {
  const sample = rows.slice(0, 30).map((r) => r[h]).filter(Boolean);
  if (!sample.length) return false;
  const n = sample.filter((v) => v instanceof Date || !isNaN(Date.parse(v))).length;
  return n / sample.length > 0.6;
}

function fillSelects() {
  const prev = { cat: catCol.value, val: valCol.value, date: dateCol.value };
  const nums = headers.filter(isNumericCol);
  const cats = headers.filter((h) => !isNumericCol(h));
  const dates = headers.filter(isDateCol);
  const opts = headers.map((h) => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`).join("");
  catCol.innerHTML = opts;
  valCol.innerHTML = (nums.length ? nums : headers).map((h) => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`).join("");
  dateCol.innerHTML = `<option value="">بدون</option>` + opts;
  catCol.value = headers.includes(prev.cat) ? prev.cat : (cats[0] || headers[0] || "");
  valCol.value = (nums.includes(prev.val) || headers.includes(prev.val)) ? prev.val : (nums[0] || headers[0] || "");
  dateCol.value = headers.includes(prev.date) ? prev.date : (dates[0] || "");
}

function num(v) {
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function filteredRows() {
  const q = search.value.trim().toLowerCase();
  if (!q) return rows;
  return rows.filter((r) => headers.some((h) => String(r[h]).toLowerCase().includes(q)));
}

function groupSum(data, key, val) {
  const map = new Map();
  data.forEach((r) => {
    const k = String(r[key] ?? "غير محدد") || "غير محدد";
    map.set(k, (map.get(k) || 0) + num(r[val]));
  });
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, 12);
}

function renderKpis(data) {
  const v = valCol.value;
  const values = data.map((r) => num(r[v]));
  const sum = values.reduce((a, b) => a + b, 0);
  const avg = values.length ? sum / values.length : 0;
  const max = values.length ? Math.max(...values) : 0;
  const min = values.length ? Math.min(...values) : 0;
  const items = [
    ["عدد الصفوف", data.length.toLocaleString("ar-EG")],
    ["المجموع", sum.toLocaleString("ar-EG", { maximumFractionDigits: 2 })],
    ["المتوسط", avg.toLocaleString("ar-EG", { maximumFractionDigits: 2 })],
    ["أعلى قيمة", max.toLocaleString("ar-EG", { maximumFractionDigits: 2 })],
    ["أقل قيمة", min.toLocaleString("ar-EG", { maximumFractionDigits: 2 })],
  ];
  document.getElementById("kpis").innerHTML = items
    .map(([label, value]) => `<div class="kpi"><div class="label">${label}</div><div class="value">${value}</div></div>`)
    .join("");
}

function destroyCharts() {
  Object.values(charts).forEach((c) => c && c.destroy());
}

const chartBase = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: { legend: { display: false } },
  scales: {
    x: { ticks: { maxRotation: 40, autoSkip: true, maxTicksLimit: 8, font: { size: 10 } } },
    y: { ticks: { maxTicksLimit: 5, font: { size: 10 } } },
  },
};

function displayCols() {
  const selected = [catCol.value, valCol.value, dateCol.value].filter((h) => h && headers.includes(h));
  const unique = [...new Set(selected)];
  if (document.getElementById("onlySelectedCols").checked && unique.length) return unique;
  return headers;
}

function formatCell(v) {
  if (v instanceof Date && !isNaN(v)) return v.toLocaleDateString("ar-EG");
  if (v == null) return "";
  return String(v);
}

function renderCharts(data) {
  destroyCharts();
  const grouped = groupSum(data, catCol.value, valCol.value);
  const labels = grouped.length ? grouped.map((x) => x[0]) : ["لا توجد بيانات"];
  const vals = grouped.length ? grouped.map((x) => x[1]) : [0];
  const colors = labels.map((_, i) => `hsl(${(i * 37) % 360} 70% 55%)`);

  charts.bar = new Chart(document.getElementById("barChart"), {
    type: "bar",
    data: { labels, datasets: [{ label: valCol.value, data: vals, backgroundColor: colors }] },
    options: { ...chartBase, plugins: { legend: { display: false } } },
  });
  charts.pie = new Chart(document.getElementById("pieChart"), {
    type: "doughnut",
    data: { labels, datasets: [{ data: vals, backgroundColor: colors }] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: true, position: "bottom", labels: { boxWidth: 8, font: { size: 10 } } } },
    },
  });

  const dCol = dateCol.value;
  let lineLabels = [];
  let lineVals = [];
  if (dCol) {
    const byDate = new Map();
    data.forEach((r) => {
      const d = r[dCol] instanceof Date ? r[dCol] : new Date(r[dCol]);
      if (isNaN(d)) return;
      const key = d.toISOString().slice(0, 10);
      byDate.set(key, (byDate.get(key) || 0) + num(r[valCol.value]));
    });
    const sorted = [...byDate.entries()].sort((a, b) => a[0].localeCompare(b[0]));
    lineLabels = sorted.map((x) => x[0]);
    lineVals = sorted.map((x) => x[1]);
  }
  charts.line = new Chart(document.getElementById("lineChart"), {
    type: "line",
    data: {
      labels: lineLabels.length ? lineLabels : ["لا يوجد عمود تاريخ"],
      datasets: [{
        label: valCol.value,
        data: lineVals.length ? lineVals : [0],
        borderColor: "#3dd6c6",
        backgroundColor: "rgba(61,214,198,.2)",
        fill: true,
        tension: 0.3,
      }],
    },
    options: { ...chartBase, plugins: { legend: { display: false } } },
  });
}

function fillRuleCol() {
  const sel = document.getElementById("ruleCol");
  sel.innerHTML = headers.map((h) => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`).join("");
}

function applyCurrentRule() {
  const col = document.getElementById("ruleCol").value;
  if (!col) return;
  colRules[col] = {
    type: document.getElementById("ruleType").value,
    min: document.getElementById("ruleMin").value.trim(),
    max: document.getElementById("ruleMax").value.trim(),
    list: document.getElementById("ruleList").value.trim(),
  };
  renderRulesList();
  renderAll();
}

function renderRulesList() {
  const box = document.getElementById("rulesList");
  const entries = Object.entries(colRules);
  box.innerHTML = entries.length
    ? entries.map(([col, r]) => `<span class="rule-tag">${escapeHtml(col)}: ${escapeHtml(r.type)}${r.min ? " ≥ " + escapeHtml(r.min) : ""}${r.max ? " ≤ " + escapeHtml(r.max) : ""}</span>`).join("")
    : '<span class="rule-tag">لا توجد قواعد مخصصة — يُستخدم التحقق التلقائي</span>';
}

function autoRule(h) {
  if (colRules[h]) return colRules[h];
  if (h === valCol.value) return { type: "number", min: "", max: "", list: "" };
  if (h === dateCol.value) return { type: "date", min: "", max: "", list: "" };
  if (h === catCol.value) return { type: "required", min: "", max: "", list: "" };
  return { type: "auto", min: "", max: "", list: "" };
}

function isEmpty(v) {
  return v === "" || v == null || (typeof v === "string" && v.trim() === "");
}

function looksNumber(v) {
  if (typeof v === "number" && Number.isFinite(v)) return true;
  return /^-?\d+(\.\d+)?$/.test(String(v).replace(/,/g, "").trim());
}

function looksDate(v) {
  if (v instanceof Date && !isNaN(v)) return true;
  const s = String(v).trim();
  if (!s) return false;
  return !isNaN(Date.parse(s));
}

function rowKey(r) {
  return headers.map((h) => String(r[h] ?? "")).join("||");
}

function inList(v, listStr) {
  const allowed = listStr.split(/[,،]/).map((s) => s.trim()).filter(Boolean);
  if (!allowed.length) return true;
  return allowed.includes(String(v).trim());
}

function inRange(v, min, max, asDate) {
  if (asDate) {
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d)) return false;
    if (min && d < new Date(min)) return false;
    if (max && d > new Date(max)) return false;
    return true;
  }
  const n = num(v);
  if (min !== "" && n < Number(min)) return false;
  if (max !== "" && n > Number(max)) return false;
  return true;
}

function validateRows(data) {
  const cols = displayCols();
  const uniqueCols = cols.filter((h) => autoRule(h).type === "unique");
  const uniqueMaps = {};
  uniqueCols.forEach((h) => {
    const m = new Map();
    data.forEach((r) => {
      const k = String(r[h] ?? "");
      m.set(k, (m.get(k) || 0) + 1);
    });
    uniqueMaps[h] = m;
  });
  const rowCounts = new Map();
  data.forEach((r) => rowCounts.set(rowKey(r), (rowCounts.get(rowKey(r)) || 0) + 1));

  let emptyCells = 0;
  let badNums = 0;
  let badDates = 0;
  let dups = 0;
  let badList = 0;
  let badRange = 0;
  const issues = data.map((r) => {
    const cells = {};
    let rowInvalid = false;
    cols.forEach((h) => {
      const rule = autoRule(h);
      const v = r[h];
      const type = rule.type === "auto" ? (isNumericCol(h) ? "number" : isDateCol(h) ? "date" : "none") : rule.type;
      if (type === "none") return;
      if (isEmpty(v)) {
        if (type === "required" || type === "number" || type === "integer" || type === "date" || type === "unique" || type === "list") {
          cells[h] = "empty";
          emptyCells += 1;
          rowInvalid = true;
        }
        return;
      }
      if ((type === "number" || type === "integer") && !looksNumber(v)) {
        cells[h] = "bad-num";
        badNums += 1;
        rowInvalid = true;
        return;
      }
      if (type === "integer" && !Number.isInteger(Number(String(v).replace(/,/g, "")))) {
        cells[h] = "bad-num";
        badNums += 1;
        rowInvalid = true;
        return;
      }
      if (type === "date" && !looksDate(v)) {
        cells[h] = "bad-date";
        badDates += 1;
        rowInvalid = true;
        return;
      }
      if ((type === "number" || type === "integer" || type === "date") && !inRange(v, rule.min, rule.max, type === "date")) {
        cells[h] = "bad-range";
        badRange += 1;
        rowInvalid = true;
        return;
      }
      if (type === "list" && !inList(v, rule.list)) {
        cells[h] = "bad-list";
        badList += 1;
        rowInvalid = true;
        return;
      }
      if (type === "unique" && uniqueMaps[h] && uniqueMaps[h].get(String(v)) > 1) {
        cells[h] = "bad-unique";
        rowInvalid = true;
      }
    });
    const dup = rowCounts.get(rowKey(r)) > 1;
    if (dup) {
      dups += 1;
      rowInvalid = true;
    }
    return { cells, dup, rowInvalid };
  });
  return { issues, emptyCells, badNums, badDates, dups, badList, badRange };
}

function renderValidation(stats, invalidCount, total) {
  const ok = invalidCount === 0;
  const chips = [
    `<span class="vchip ${ok ? "ok" : "err"}">${ok ? "البيانات سليمة" : `${invalidCount} صف به مشاكل`} من ${total}</span>`,
    `<span class="vchip ${stats.emptyCells ? "warn" : "ok"}">خلايا فارغة: ${stats.emptyCells}</span>`,
    `<span class="vchip ${stats.badNums ? "err" : "ok"}">أرقام غير صالحة: ${stats.badNums}</span>`,
    `<span class="vchip ${stats.badDates ? "err" : "ok"}">تواريخ غير صالحة: ${stats.badDates}</span>`,
    `<span class="vchip ${stats.dups ? "warn" : "ok"}">صفوف مكررة: ${stats.dups}</span>`,
  ];
  document.getElementById("validationBar").innerHTML = chips.join("");
}

function renderTable(data) {
  const { issues, emptyCells, badNums, badDates, dups } = validateRows(data);
  const onlyInvalid = document.getElementById("onlyInvalid").checked;
  const invalidCount = issues.filter((i) => i.rowInvalid).length;
  renderValidation({ emptyCells, badNums, badDates, dups }, invalidCount, data.length);

  const thead = `<tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr>`;
  let shown = 0;
  const tbody = data
    .map((r, idx) => {
      const issue = issues[idx];
      if (onlyInvalid && !issue.rowInvalid) return "";
      if (shown >= 200) return "";
      shown += 1;
      const cls = issue.rowInvalid ? "invalid" : "";
      const tds = headers
        .map((h) => {
          const kind = issue.cells[h] || (issue.dup ? "dup" : "");
          const val = r[h] instanceof Date ? r[h].toLocaleDateString("ar-EG") : r[h];
          const title = kind === "empty" ? "خلية فارغة" : kind === "bad-num" ? "قيمة رقمية غير صالحة" : kind === "bad-date" ? "تاريخ غير صالح" : issue.dup ? "صف مكرر" : "";
          return `<td class="${kind}" title="${title}">${val ?? ""}</td>`;
        })
        .join("");
      return `<tr class="${cls}">${tds}</tr>`;
    })
    .join("");
  document.getElementById("dataTable").innerHTML = thead + tbody;
}

function renderAll() {
  const data = filteredRows();
  renderKpis(data);
  renderCharts(data);
  renderTable(data);
}
