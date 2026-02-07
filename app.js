// Monthly Meals Counter
// Rules:
// - Sum columns I,J,K,L (default) per row.
// - If column I is "A" or "F" => absent => 0 meals for that row.
// - If "PB" appears anywhere in meal columns (case-insensitive) => 0 meals for that row.
// - If person has only PB days across all rows => show "PB" instead of a number.
// - Names matched case-insensitively, accent-insensitively, and ignoring special characters like '*'.

const fileInput = document.getElementById("fileInput");
const runBtn = document.getElementById("runBtn");
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

fileInput.addEventListener("change", () => {
  runBtn.disabled = !(fileInput.files && fileInput.files.length);
  downloadCsvBtn.disabled = true;
  downloadXlsxBtn.disabled = true;
  searchInput.disabled = true;
  sortSelect.disabled = true;
  resultsBody.innerHTML = "";
  status("");
});

runBtn.addEventListener("click", async () => {
  try {
    status("Reading files…");
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

    const accumulator = new Map(); // key: normalizedName => stats

    for (const f of files) {
      status(`Parsing: ${f.name}`);
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;

        // 2D array; empty cells become null
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
        if (!rows || !rows.length) continue;

        const startRow = forcedFirstRow > 0
            ? (forcedFirstRow - 1)
            : autoDetectFirstDataRow(rows, nameColIdx, mealColIdxs);

        for (let r = startRow; r < rows.length; r++) {
          const row = rows[r] || [];
          const rawName = row[nameColIdx];

          // Cleaned name for DISPLAY (removes '*' etc, but keeps accents)
          const displayName = normalizeDisplayName(rawName);
          if (!displayName) continue;

          // Key ignores accents, case, and special chars like '*'
          const key = normalizeKey(displayName);

          // Determine row status & meals
          const colI = row[mealColIdxs[0]]; // first meal col (default I)
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
              name: displayName,   // CLEANED display name stored
              meals: 0,
              pbDays: 0,
              mealDays: 0,
              rows: 0,
              everNumericMeals: false,
              everPB: false
            });
          } else {
            // Prefer a "better" display name if we find one later
            const stExisting = accumulator.get(key);
            stExisting.name = chooseBetterDisplayName(stExisting.name, displayName);
          }

          const st = accumulator.get(key);
          st.rows += 1;

          if (pbDay) {
            st.pbDays += 1;
            st.everPB = true;
            continue; // PB day always counts as 0 meals
          }

          if (absent) {
            continue; // A/F day counts as 0 meals
          }

          st.meals += rowMeals;

          if (mealDay) {
            st.mealDays += 1;
            st.everNumericMeals = true;
          }
        }
      }
    }

    // Build report rows
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

    // Enable UI
    const hasResults = latestRows.length > 0;
    downloadCsvBtn.disabled = !hasResults;
    downloadXlsxBtn.disabled = !hasResults;
    searchInput.disabled = !hasResults;
    sortSelect.disabled = !hasResults;

    render();
    status(`Done. Processed ${files.length} file(s). People found: ${latestRows.length}.`);
  } catch (err) {
    console.error(err);
    status(`Error: ${err?.message || err}`);
  }
});

searchInput.addEventListener("input", render);
sortSelect.addEventListener("change", render);

// CSV export
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

// Excel (.xlsx) export
downloadXlsxBtn.addEventListener("click", () => {
  if (!latestRows.length) return;

  const data = latestRows.map(r => ({
    Name: r.name,
    Total: r.total,          // number or "PB"
    "PB days": r.pbDays,
    "Meal days": r.mealDays,
    "Rows counted": r.rows
  }));

  const ws = XLSX.utils.json_to_sheet(data, { skipHeader: false });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Monthly Report");

  // Basic column widths
  ws["!cols"] = [
    { wch: 35 }, // Name
    { wch: 10 }, // Total
    { wch: 10 }, // PB days
    { wch: 10 }, // Meal days
    { wch: 14 }  // Rows counted
  ];

  XLSX.writeFile(wb, "monthly_meals_report.xlsx");
});

// ---------------- Helpers ----------------

function status(msg) {
  statusEl.textContent = msg || "";
}

function colLetterToIndex(letter) {
  // A => 0, B => 1, ..., Z => 25, AA => 26
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
  if (typeof v === "number") return ""; // names shouldn't be numeric
  let s = String(v).trim();
  if (!s) return "";

  // Normalize unicode (keeps accents), remove special chars like '*' but keep letters/numbers/spaces
  s = s.normalize("NFKC")
      .replace(/[^\p{L}\p{N}\s]/gu, "") // strips *, punctuation, etc (keeps accented letters)
      .replace(/\s+/g, " ")
      .trim();

  if (!s) return "";

  // avoid common header junk (compare using key-normalized)
  const sKey = normalizeKey(s);
  if (["familias", "families", "family", "nome", "name"].includes(sKey)) return "";

  return s;
}

// Key normalization: ignores accents + case + special characters (used ONLY for merging/searching)
function normalizeKey(s) {
  return String(s)
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "") // remove accents/diacritics
      .replace(/[^a-z0-9\s]/g, "")     // remove special characters like * (keeps letters/numbers/spaces)
      .replace(/\s+/g, " ");           // collapse spaces
}

// If we see the same person written differently, pick the nicer display name:
// - prefer the one with diacritics (José vs Jose)
// - otherwise prefer longer (often includes middle/last name)
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

    // handle numeric strings like "10" or " 2 "
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
    if (r > 3) return r; // usually below header banner
  }
  return 0;
}

function render() {
  const q = normalizeKey(searchInput.value || "");
  let rows = latestRows.slice();

  if (q) {
    rows = rows.filter(r => normalizeKey(r.name).includes(q));
  }

  const sortMode = sortSelect.value;
  rows.sort((a, b) => {
    if (sortMode === "nameAsc") return normalizeKey(a.name).localeCompare(normalizeKey(b.name));
    if (sortMode === "nameDesc") return normalizeKey(b.name).localeCompare(normalizeKey(a.name));

    // meals sort: PB treated as -1
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
    const vals = [
      csvCell(r.name),
      csvCell(String(r.total)),
      r.pbDays,
      r.mealDays,
      r.rows
    ];
    lines.push(vals.join(","));
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
