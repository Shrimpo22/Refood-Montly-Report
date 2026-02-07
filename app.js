// Monthly Meals Counter + Fill Month Report Template

const fileInput = document.getElementById("fileInput");
const reportInput = document.getElementById("reportInput");

const runBtn = document.getElementById("runBtn");
const fillReportBtn = document.getElementById("fillReportBtn");

const statusEl = document.getElementById("status");
const resultsBody = document.querySelector("#resultsTable tbody");

const downloadCsvBtn = document.getElementById("downloadCsvBtn");
const downloadXlsxBtn = document.getElementById("downloadXlsxBtn");

const searchInput = document.getElementById("search");
const sortSelect = document.getElementById("sort");

const nameColInput = document.getElementById("nameCol");
const mealColsInput = document.getElementById("mealCols");
const firstRowInput = document.getElementById("firstRow");

let latestRows = []; // for rendering + exports
let latestMap = new Map(); // key -> { name, total } (total is number or "PB")

let reportFile = null;

fileInput.addEventListener("change", () => {
  runBtn.disabled = !(fileInput.files && fileInput.files.length);

  downloadCsvBtn.disabled = true;
  downloadXlsxBtn.disabled = true;
  searchInput.disabled = true;
  sortSelect.disabled = true;

  latestRows = [];
  latestMap = new Map();
  resultsBody.innerHTML = "";
  status("");

  updateFillButtonState();
});

reportInput.addEventListener("change", () => {
  reportFile = (reportInput.files && reportInput.files[0]) ? reportInput.files[0] : null;
  updateFillButtonState();
});

function updateFillButtonState() {
  const hasMeals = latestRows.length > 0;
  const hasReport = !!reportFile;
  fillReportBtn.disabled = !(hasMeals && hasReport);
}

runBtn.addEventListener("click", async () => {
  try {
    status("Reading meal files…");
    const files = Array.from(fileInput.files || []);
    if (!files.length) return;

    const nameColLetter = (nameColInput.value || "E").trim().toUpperCase();
    const mealColLetters = (mealColsInput.value || "I,J,K,L")
        .split(",")
        .map(s => s.trim().toUpperCase())
        .filter(Boolean);

    const forcedFirstRow = parseInt(firstRowInput.value || "0", 10) || 0;

    const nameColIdx = colLetterToIndex(nameColLetter);
    const mealColIdxs = mealColLetters.map(colLetterToIndex);

    const accumulator = new Map(); // key -> stats

    for (const f of files) {
      status(`Parsing: ${f.name}`);
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;

        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
        if (!rows || !rows.length) continue;

        const startRow = forcedFirstRow > 0
            ? (forcedFirstRow - 1)
            : autoDetectFirstDataRow(rows, nameColIdx, mealColIdxs);

        for (let r = startRow; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawName = row[nameColIdx];

          const displayName = normalizeDisplayName(rawName);
          if (!displayName) continue;

          const key = normalizeKey(displayName);

          const colI = row[mealColIdxs[0]];
          const absent = isAbsent(colI);
          const pbDay = isPB(colI) || anyPB(row, mealColIdxs);

          let rowMeals = 0;
          let mealDay = false;

          if (!absent && !pbDay) {
            rowMeals = sumNumeric(row, mealColIdxs);
            if (rowMeals > 0) mealDay = true;
          }

          if (!accumulator.has(key)) {
            accumulator.set(key, {
              name: displayName,
              meals: 0,
              pbDays: 0,
              mealDays: 0,
              rows: 0,
              everNumericMeals: false,
              everPB: false
            });
          } else {
            const stExisting = accumulator.get(key);
            stExisting.name = chooseBetterDisplayName(stExisting.name, displayName);
          }

          const st = accumulator.get(key);
          st.rows += 1;

          if (pbDay) {
            st.pbDays += 1;
            st.everPB = true;
            continue;
          }
          if (absent) continue;

          st.meals += rowMeals;

          if (mealDay) {
            st.mealDays += 1;
            st.everNumericMeals = true;
          }
        }
      }
    }

    latestRows = Array.from(accumulator.values()).map(st => {
      const onlyPB = st.everPB && !st.everNumericMeals && st.meals === 0;
      return {
        name: st.name,
        total: onlyPB ? "PB" : st.meals,
        mealsNumeric: onlyPB ? null : st.meals,
        pbDays: st.pbDays,
        mealDays: st.mealDays,
        rows: st.rows
      };
    });

    // Build lookup map for filling template
    latestMap = new Map(
        latestRows.map(r => [normalizeKey(r.name), { name: r.name, total: r.total }])
    );

    const hasResults = latestRows.length > 0;
    downloadCsvBtn.disabled = !hasResults;
    downloadXlsxBtn.disabled = !hasResults;
    searchInput.disabled = !hasResults;
    sortSelect.disabled = !hasResults;

    render();
    status(`Done. Processed ${files.length} meal file(s). People found: ${latestRows.length}.`);

    updateFillButtonState();
  } catch (err) {
    console.error(err);
    status(`Error: ${err?.message || err}`);
  }
});

fillReportBtn.addEventListener("click", async () => {
  if (!reportFile || latestRows.length === 0) return;

  try {
    status(`Filling month report: ${reportFile.name}`);
    const buf = await reportFile.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const targetSheetName = findBeneficiariosSheet(wb);
    if (!targetSheetName) {
      status(`Error: couldn't find a "Beneficiários" sheet in the report.`);
      return;
    }

    const ws = wb.Sheets[targetSheetName];

    const NAME_COL = "C";
    const FAMILY_COL = "D";
    const MEALS_COL = "Q";

    const headerRow = findHeaderRow(ws, NAME_COL);
    const startRow = headerRow ? (headerRow + 1) : 8; // fallback to 8

    // Collect existing names in template and rewrite Q
    const existingKeys = new Set();
    let lastRow = startRow;

    for (let row = startRow; ; row++) {
      const nameCellAddr = `${NAME_COL}${row}`;
      const nameCell = ws[nameCellAddr];
      const rawName = nameCell ? nameCell.v : null;

      const disp = normalizeDisplayName(rawName);
      if (!disp) break; // stop at first empty name row (template uses contiguous list)

      const key = normalizeKey(disp);
      existingKeys.add(key);
      lastRow = row;

      const match = latestMap.get(key);
      const newValue = match ? match.total : 0;

      setCell(ws, `${MEALS_COL}${row}`, newValue);
    }

    // Append missing people
    const missing = [];
    for (const [key, info] of latestMap.entries()) {
      if (!existingKeys.has(key)) missing.push(info);
    }

    let writeRow = lastRow + 1;
    for (const info of missing.sort((a, b) => normalizeKey(a.name).localeCompare(normalizeKey(b.name)))) {
      setCell(ws, `${NAME_COL}${writeRow}`, info.name);
      setCell(ws, `${FAMILY_COL}${writeRow}`, 1);
      setCell(ws, `${MEALS_COL}${writeRow}`, info.total);
      writeRow++;
    }

    XLSX.writeFile(wb, `FILLED_${reportFile.name.replace(/\.(xlsx|xls)$/i, "")}.xlsx`);
    status(`Filled report downloaded. Updated ${targetSheetName}: Q column rewritten, missing people appended.`);
  } catch (err) {
    console.error(err);
    status(`Error filling report: ${err?.message || err}`);
  }
});

// ----------- Exports you already had -----------

searchInput.addEventListener("input", render);
sortSelect.addEventListener("change", render);

downloadCsvBtn.addEventListener("click", () => {
  if (!latestRows.length) return;
  const csv = toCSV(latestRows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "monthly_meals_report.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

downloadXlsxBtn.addEventListener("click", () => {
  if (!latestRows.length) return;

  const data = latestRows.map(r => ({
    Name: r.name,
    Total: r.total,
    "PB days": r.pbDays,
    "Meal days": r.mealDays,
    "Rows counted": r.rows
  }));

  const ws = XLSX.utils.json_to_sheet(data, { skipHeader: false });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Monthly Report");

  ws["!cols"] = [
    { wch: 35 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 14 }
  ];

  XLSX.writeFile(wb, "monthly_meals_report.xlsx");
});

// ---------------- Helpers ----------------

function status(msg) {
  statusEl.textContent = msg || "";
}

function colLetterToIndex(letter) {
  let s = (letter || "").toUpperCase().trim();
  if (!s) return 0;
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) continue;
    n = n * 26 + (code - 64);
  }
  return Math.max(0, n - 1);
}

function normalizeDisplayName(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return "";
  let s = String(v).trim();
  if (!s) return "";

  // keep accents, remove punctuation like **, etc (modern browsers)
  s = s.normalize("NFKC")
      .replace(/[^\p{L}\p{N}\s]/gu, "")
      .replace(/\s+/g, " ")
      .trim();

  if (!s) return "";

  const sKey = normalizeKey(s);
  if (["familias", "families", "family", "nome", "name"].includes(sKey)) return "";
  return s;
}

function normalizeKey(s) {
  return String(s)
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9\s]/g, "")
      .replace(/\s+/g, " ");
}

function chooseBetterDisplayName(current, candidate) {
  const cur = String(current || "").trim();
  const cand = String(candidate || "").trim();
  if (!cur) return cand;
  if (!cand) return cur;

  const curHasDia = hasDiacritics(cur);
  const candHasDia = hasDiacritics(cand);

  if (!curHasDia && candHasDia) return cand;
  if (curHasDia && !candHasDia) return cur;

  if (cand.length > cur.length) return cand;
  return cur;
}

function hasDiacritics(str) {
  return /[\u0300-\u036f]/.test(str.normalize("NFD"));
}

function isAbsent(v) {
  if (v === null || v === undefined) return false;
  const s = normalizeKey(v);
  return s === "a" || s === "f";
}

function isPB(v) {
  if (v === null || v === undefined) return false;
  const s = normalizeKey(v);
  return s === "pb";
}

function anyPB(row, idxs) {
  for (const i of idxs) {
    if (isPB(row[i])) return true;
  }
  return false;
}

function sumNumeric(row, idxs) {
  let total = 0;
  for (const i of idxs) {
    const v = row[i];
    if (v === null || v === undefined) continue;

    if (typeof v === "number" && Number.isFinite(v)) {
      total += v;
      continue;
    }

    const s = String(v).trim().replace(",", ".");
    const num = Number(s);
    if (Number.isFinite(num)) total += num;
  }
  return total;
}

function autoDetectFirstDataRow(rows, nameColIdx, mealColIdxs) {
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const name = normalizeDisplayName(row[nameColIdx]);
    if (!name) continue;

    let interesting = false;
    for (const c of mealColIdxs) {
      const v = row[c];
      if (v === null || v === undefined) continue;
      interesting = true;
      break;
    }

    if (interesting) return r;
    if (r > 3) return r;
  }
  return 0;
}

function render() {
  const q = normalizeKey(searchInput.value || "");
  let rows = latestRows.slice();

  if (q) rows = rows.filter(r => normalizeKey(r.name).includes(q));

  const sortMode = sortSelect.value;
  rows.sort((a, b) => {
    if (sortMode === "nameAsc") return normalizeKey(a.name).localeCompare(normalizeKey(b.name));
    if (sortMode === "nameDesc") return normalizeKey(b.name).localeCompare(normalizeKey(a.name));

    const am = a.mealsNumeric === null ? -1 : a.mealsNumeric;
    const bm = b.mealsNumeric === null ? -1 : b.mealsNumeric;

    if (sortMode === "mealsAsc") return am - bm;
    if (sortMode === "mealsDesc") return bm - am;
    return 0;
  });

  resultsBody.innerHTML = rows.map(r => {
    const total = (r.total === "PB") ? "PB" : String(r.total);
    return `
      <tr>
        <td>${escapeHTML(r.name)}</td>
        <td class="num">${escapeHTML(total)}</td>
        <td class="num">${r.pbDays}</td>
        <td class="num">${r.mealDays}</td>
        <td class="num">${r.rows}</td>
      </tr>
    `;
  }).join("");
}

function toCSV(rows) {
  const header = ["Name", "Total", "PB days", "Meal days", "Rows counted"];
  const lines = [header.join(",")];
  for (const r of rows) {
    lines.push([
      csvCell(r.name),
      csvCell(String(r.total)),
      r.pbDays,
      r.mealDays,
      r.rows
    ].join(","));
  }
  return lines.join("\n");
}

function csvCell(v) {
  const s = String(v ?? "");
  if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function escapeHTML(s) {
  return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
}

// ----------- Template-filling helpers -----------

function findBeneficiariosSheet(wb) {
  // Prefer sheets whose name contains "benef"
  const names = wb.SheetNames || [];
  for (const n of names) {
    const k = normalizeKey(n);
    if (k.includes("benef")) return n;
  }
  // Fallback: try find a sheet that has "Nome Beneficiário" in column C
  for (const n of names) {
    const ws = wb.Sheets[n];
    if (ws && findHeaderRow(ws, "C")) return n;
  }
  return null;
}

function findHeaderRow(ws, nameColLetter) {
  // scan first 60 rows in column C for "Nome Beneficiário" (normalized)
  for (let r = 1; r <= 60; r++) {
    const addr = `${nameColLetter}${r}`;
    const cell = ws[addr];
    const v = cell ? cell.v : null;
    const k = normalizeKey(v || "");
    if (k.includes("nome beneficiario")) return r;
  }
  return null;
}

function setCell(ws, addr, value) {
  // set value while ensuring cell exists
  const t = (typeof value === "number") ? "n" : "s";
  ws[addr] = { t, v: value };
  // expand range
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  const { c, r } = XLSX.utils.decode_cell(addr);
  if (r < range.s.r) range.s.r = r;
  if (c < range.s.c) range.s.c = c;
  if (r > range.e.r) range.e.r = r;
  if (c > range.e.c) range.e.c = c;
  ws["!ref"] = XLSX.utils.encode_range(range);
}
