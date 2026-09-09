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

function useSheet(name) {
  const sheet = workbook.Sheets[name];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  rows = json;
  headers = json.length ? Object.keys(json[0]) : [];
  document.getElementById("rowCount").textContent = `${rows.length} صف`;
  document.getElementById("colCount").textContent = `${headers.length} عمود`;
  fillSelects();
  renderAll();
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
  const nums = headers.filter(isNumericCol);
  const cats = headers.filter((h) => !isNumericCol(h));
  const dates = headers.filter(isDateCol);
  catCol.innerHTML = headers.map((h) => `<option>${h}</option>`).join("");
  valCol.innerHTML = (nums.length ? nums : headers).map((h) => `<option>${h}</option>`).join("");
  dateCol.innerHTML = `<option value="">بدون</option>` + headers.map((h) => `<option>${h}</option>`).join("");
  if (cats[0]) catCol.value = cats[0];
  if (nums[0]) valCol.value = nums[0];
  if (dates[0]) dateCol.value = dates[0];
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

function renderCharts(data) {
  destroyCharts();
  const grouped = groupSum(data, catCol.value, valCol.value);
  const labels = grouped.map((x) => x[0]);
  const vals = grouped.map((x) => x[1]);
  const colors = labels.map((_, i) => `hsl(${(i * 37) % 360} 70% 55%)`);

  charts.bar = new Chart(document.getElementById("barChart"), {
    type: "bar",
    data: { labels, datasets: [{ label: valCol.value, data: vals, backgroundColor: colors }] },
    options: { responsive: true, plugins: { legend: { display: false } } },
  });
  charts.pie = new Chart(document.getElementById("pieChart"), {
    type: "doughnut",
    data: { labels, datasets: [{ data: vals, backgroundColor: colors }] },
    options: { responsive: true },
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
    options: { responsive: true },
  });
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

function validateRows(data) {
  const numericCols = headers.filter(isNumericCol);
  const dateCols = headers.filter(isDateCol);
  const counts = new Map();
  data.forEach((r) => counts.set(rowKey(r), (counts.get(rowKey(r)) || 0) + 1));

  let emptyCells = 0;
  let badNums = 0;
  let badDates = 0;
  let dups = 0;
  const issues = data.map((r) => {
    const cells = {};
    let rowInvalid = false;
    headers.forEach((h) => {
      const v = r[h];
      if (isEmpty(v)) {
        cells[h] = "empty";
        emptyCells += 1;
        rowInvalid = true;
        return;
      }
      if (numericCols.includes(h) && !looksNumber(v)) {
        cells[h] = "bad-num";
        badNums += 1;
        rowInvalid = true;
      }
      if (dateCols.includes(h) && !looksDate(v)) {
        cells[h] = "bad-date";
        badDates += 1;
        rowInvalid = true;
      }
    });
    const dup = counts.get(rowKey(r)) > 1;
    if (dup) {
      dups += 1;
      rowInvalid = true;
    }
    return { cells, dup, rowInvalid };
  });
  return { issues, emptyCells, badNums, badDates, dups };
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
