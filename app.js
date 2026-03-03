/**
 * System DI Smart Search
 * Fetch Google Sheet range D10:N974 (sheet DB_DI_DUC)
 * Map columns (range starts at D):
 *   D=Name (idx 0)
 *   F=Gender (idx 2)
 *   G=Class (idx 3)
 *   H=Generation (idx 4)
 *   J=Role (idx 6)
 *   K=Status (idx 7) ✅
 *   M=ID (idx 9)
 *   N=Group (idx 10)
 */

const SHEET_ID = "1h6pqlcoUSKPWsk7it4jFLtg5Oc_w7gXGnKCXSIduK7E";
const GID = "864623093";
const SHEET_NAME = "DB_DI_DUC";
const RANGE = "D10:N974"; // ✅ includes J,K,M,N

const GVIZ_URL =
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq` +
  `?gid=${encodeURIComponent(GID)}` +
  `&sheet=${encodeURIComponent(SHEET_NAME)}` +
  `&range=${encodeURIComponent(RANGE)}` +
  `&tqx=out:json`;

const els = {
  qName: document.getElementById("qName"),
  qId: document.getElementById("qId"),
  qStatus: document.getElementById("qStatus"),
  qRole: document.getElementById("qRole"),

  btnReload: document.getElementById("btnReload"),
  btnClear: document.getElementById("btnClear"),
  btnPrint: document.getElementById("btnPrint"),
  tbody: document.getElementById("tbody"),

  kpiStaff: document.getElementById("kpiStaff"),
  kpiMale: document.getElementById("kpiMale"),
  kpiFemale: document.getElementById("kpiFemale"),

  statusDot: document.getElementById("statusDot"),
  statusText: document.getElementById("statusText"),

  btnTheme: document.getElementById("btnTheme"),
  themeText: document.getElementById("themeText"),
  themeIcon: document.querySelector(".toggleIcon"),

  // group dropdown
  groupDD: document.getElementById("groupDD"),
  btnGroupDD: document.getElementById("btnGroupDD"),
  groupMenu: document.getElementById("groupMenu"),
  groupList: document.getElementById("groupList"),
  groupDDText: document.getElementById("groupDDText"),
  groupSearch: document.getElementById("groupSearch"),
  btnGroupAll: document.getElementById("btnGroupAll"),
  btnGroupNone: document.getElementById("btnGroupNone"),
};

let CACHE = { rows: [] };
let LAST_FILTERED = [];
let GROUPS = [];
let GROUP_SELECTED = new Set();

/* =========================
   THEME (Dark/Light)
========================= */
function applyTheme(theme) {
  document.body.setAttribute("data-theme", theme);
  localStorage.setItem("theme", theme);

  const isDark = theme === "dark";
  els.themeText.textContent = isDark ? "Dark" : "Light";
  els.themeIcon.textContent = isDark ? "🌙" : "☀️";
}

(function initTheme() {
  const saved = localStorage.getItem("theme");
  applyTheme(saved === "light" ? "light" : "dark");
})();

els.btnTheme.addEventListener("click", () => {
  const cur = document.body.getAttribute("data-theme");
  applyTheme(cur === "dark" ? "light" : "dark");
});

/* =========================
   Helpers
========================= */
function setStatus(text, mode = "idle") {
  els.statusText.textContent = text;

  if (mode === "ok") {
    els.statusDot.style.background = "var(--ok)";
    els.statusDot.style.boxShadow = "0 0 0 4px rgba(47,211,155,.16)";
    return;
  }
  if (mode === "busy") {
    els.statusDot.style.background = "var(--warn)";
    els.statusDot.style.boxShadow = "0 0 0 4px rgba(245,194,91,.16)";
    return;
  }
  if (mode === "err") {
    els.statusDot.style.background = "var(--err)";
    els.statusDot.style.boxShadow = "0 0 0 4px rgba(255,107,139,.16)";
    return;
  }
  els.statusDot.style.background = "var(--accent)";
  els.statusDot.style.boxShadow = "0 0 0 4px rgba(106,183,255,.16)";
}

function debounce(fn, ms = 180) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

function norm(s) {
  return (s ?? "").toString().trim().replace(/\s+/g, " ").toLowerCase();
}

function cellToText(c) {
  if (!c) return "";
  if (typeof c.v === "string" || typeof c.v === "number") return String(c.v);
  if (c.f) return String(c.f);
  return "";
}

function escapeHtml(s) {
  return (s ?? "")
    .toString()
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// detect male/female (supports: ប្រុស/ស្រី, Male/Female, M/F, ប/ស)
function genderType(g) {
  const x = norm(g);
  if (!x) return "unknown";
  if (x.includes("ប្រុស") || x === "m" || x.includes("male") || x === "ប") return "male";
  if (x.includes("ស្រី") || x === "f" || x.includes("female") || x === "ស") return "female";
  return "unknown";
}

function countMaleFemale(rows) {
  let male = 0, female = 0;
  for (const r of rows) {
    const t = genderType(r.gender);
    if (t === "male") male++;
    else if (t === "female") female++;
  }
  return { male, female };
}

/* =========================
   Group Dropdown UI
========================= */
function openGroupDD(open) {
  const isOpen = open ?? !els.groupDD.classList.contains("open");
  els.groupDD.classList.toggle("open", isOpen);
  els.btnGroupDD.setAttribute("aria-expanded", String(isOpen));
  if (isOpen) {
    els.groupSearch.value = "";
    renderGroupList();
    setTimeout(() => els.groupSearch.focus(), 30);
  }
}

function updateGroupDDText() {
  const count = GROUP_SELECTED.size;
  els.groupDDText.textContent = count ? `បានជ្រើស ${count} ក្រុម` : "ក្រុមទាំងអស់";
}

function renderGroupList() {
  const q = norm(els.groupSearch.value);
  const list = q ? GROUPS.filter(g => norm(g).includes(q)) : GROUPS;

  els.groupList.innerHTML = list.map(g => {
    const key = norm(g);
    const checked = GROUP_SELECTED.has(key) ? "checked" : "";
    return `
      <label class="ddItem">
        <input type="checkbox" data-group="${escapeHtml(g)}" ${checked}>
        <span>${escapeHtml(g)}</span>
      </label>
    `;
  }).join("");
}

function setAllGroupsSelected(on) {
  GROUP_SELECTED.clear();
  if (on) {
    for (const g of GROUPS) GROUP_SELECTED.add(norm(g));
  }
  updateGroupDDText();
  renderGroupList();
  render();
}

/* =========================
   Status + Role options
========================= */
function buildStatusOptions(rows) {
  const current = norm(els.qStatus.value);
  const statuses = Array.from(new Set(rows.map(r => r.status).filter(Boolean)))
    .sort((a,b) => a.localeCompare(b));

  els.qStatus.innerHTML =
    `<option value="">ទាំងអស់</option>` +
    statuses.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");

  if (current) {
    const found = statuses.find(s => norm(s) === current);
    if (found) els.qStatus.value = found;
  }
}

function buildRoleOptions(rows) {
  const current = norm(els.qRole.value);
  const roles = Array.from(new Set(rows.map(r => r.role).filter(Boolean)))
    .sort((a,b) => a.localeCompare(b));

  els.qRole.innerHTML =
    `<option value="">ទាំងអស់</option>` +
    roles.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");

  if (current) {
    const found = roles.find(s => norm(s) === current);
    if (found) els.qRole.value = found;
  }
}

/* =========================
   Data
========================= */
async function fetchSheet() {
  setStatus("Loading…", "busy");

  const res = await fetch(GVIZ_URL, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);

  const text = await res.text();
  const jsonStr = text.replace(/^[\s\S]*setResponse\(/, "").replace(/\);\s*$/, "");
  const data = JSON.parse(jsonStr);

  const rows = data?.table?.rows ?? [];

  const mapped = rows.map((r) => {
    const c = r.c || [];
    return {
      nameKh: cellToText(c[0]),   // D
      gender: cellToText(c[2]),   // F
      clazz:  cellToText(c[3]),   // G
      gen:    cellToText(c[4]),   // H
      role:   cellToText(c[6]),   // J
      status: cellToText(c[7]),   // K ✅
      id:     cellToText(c[9]),   // M
      group:  cellToText(c[10]),  // N
    };
  }).filter(r => r.nameKh || r.id || r.group || r.gender || r.clazz || r.gen || r.role || r.status);

  CACHE.rows = mapped;

  els.kpiStaff.textContent = String(mapped.length);

  GROUPS = Array.from(new Set(mapped.map(r => r.group).filter(Boolean)))
    .sort((a,b) => a.localeCompare(b));

  // remove missing group selections
  for (const k of Array.from(GROUP_SELECTED)) {
    if (!GROUPS.some(g => norm(g) === k)) GROUP_SELECTED.delete(k);
  }
  updateGroupDDText();
  renderGroupList();

  buildStatusOptions(mapped);
  buildRoleOptions(mapped);

  render();
  setStatus("Loaded ✓", "ok");
}

/* =========================
   Filtering
========================= */
function applyFilter(rows) {
  const qName = norm(els.qName.value);
  const qId = norm(els.qId.value);
  const qStatus = norm(els.qStatus.value);
  const qRole = norm(els.qRole.value);

  return rows.filter(r => {
    const n = norm(r.nameKh);
    const id = norm(r.id);
    const g = norm(r.group);
    const st = norm(r.status);
    const rl = norm(r.role);

    if (qName && !n.includes(qName)) return false;
    if (qId && !id.includes(qId)) return false;
    if (GROUP_SELECTED.size && !GROUP_SELECTED.has(g)) return false;
    if (qStatus && st !== qStatus) return false;
    if (qRole && rl !== qRole) return false;

    return true;
  });
}

function render() {
  const rows = CACHE.rows || [];
  const filtered = applyFilter(rows);
  LAST_FILTERED = filtered;

  const mf = countMaleFemale(filtered);
  els.kpiMale.textContent = String(mf.male);
  els.kpiFemale.textContent = String(mf.female);

  if (!rows.length) {
    els.tbody.innerHTML = `
      <tr class="empty"><td colspan="9">មិនទាន់មានទិន្នន័យ… ចុច Reload</td></tr>`;
    return;
  }
  if (!filtered.length) {
    els.tbody.innerHTML = `
      <tr class="empty"><td colspan="9">រកមិនឃើញទិន្នន័យតាម Filter នេះទេ</td></tr>`;
    return;
  }

  const LIMIT = 500;
  const list = filtered.slice(0, LIMIT);

  els.tbody.innerHTML = list.map((r, i) => `
    <tr>
      <td><span class="badge">${i + 1}</span></td>
      <td>${escapeHtml(r.nameKh)}</td>
      <td><span class="badge">${escapeHtml(r.gender || "—")}</span></td>
      <td>${escapeHtml(r.clazz || "—")}</td>
      <td>${escapeHtml(r.gen || "—")}</td>
      <td><span class="badge">${escapeHtml(r.role || "—")}</span></td>
      <td><span class="badge">${escapeHtml(r.id || "—")}</span></td>
      <td><span class="badge">${escapeHtml(r.group || "—")}</span></td>
      <td><span class="badge">${escapeHtml(r.status || "—")}</span></td>
    </tr>
  `).join("");

  if (filtered.length > LIMIT) {
    els.tbody.insertAdjacentHTML("beforeend", `
      <tr class="empty"><td colspan="9">
        បង្ហាញតែ ${LIMIT} row ដំបូង (Matched សរុប ${filtered.length})
      </td></tr>`);
  }
}

/* =========================
   Print PDF (Auto)
========================= */
function hasActiveFilter() {
  const name = (els.qName?.value || "").trim();
  const id = (els.qId?.value || "").trim();
  const status = (els.qStatus?.value || "").trim();
  const role = (els.qRole?.value || "").trim();
  return !!(name || id || status || role || GROUP_SELECTED.size);
}

function preparePrint(rows) {
  const title = document.querySelector(".brand h1")?.textContent?.trim() || "System DI";
  const sub = document.querySelector(".sub")?.textContent?.trim() || "";
  const mf = countMaleFemale(rows);

  const headerHtml = `
    <div id="printHeader" style="margin-bottom:10px">
      <h2 style="margin:0 0 4px 0;font-size:18px">Digital Industry — Report</h2>
      <div style="color:#444;font-size:12px;margin-bottom:8px">${escapeHtml(sub)}</div>
      <div style="font-size:12px;color:#111">
        បុគ្គលិកសរុប: <b>${rows.length}</b> —
        ប្រុស: <b>${mf.male}</b> —
        ស្រី: <b>${mf.female}</b>
      </div>
      <hr style="border:none;border-top:1px solid #ddd;margin:10px 0 0 0">
    </div>
  `;

  const panel = document.querySelector(".panel");
  if (panel && !document.getElementById("printHeader")) {
    panel.insertAdjacentHTML("afterbegin", headerHtml);
  }

  const tbody = document.getElementById("tbody");
  tbody.setAttribute("data-original-html", tbody.innerHTML);

  tbody.innerHTML = rows.map((r, i) => `
    <tr>
      <td>${i + 1}</td>
      <td>${escapeHtml(r.nameKh)}</td>
      <td>${escapeHtml(r.gender || "—")}</td>
      <td>${escapeHtml(r.clazz || "—")}</td>
      <td>${escapeHtml(r.gen || "—")}</td>
      <td>${escapeHtml(r.role || "—")}</td>
      <td>${escapeHtml(r.id || "—")}</td>
      <td>${escapeHtml(r.group || "—")}</td>
      <td>${escapeHtml(r.status || "—")}</td>
    </tr>
  `).join("");
}

function cleanupPrint() {
  const tbody = document.getElementById("tbody");
  const original = tbody.getAttribute("data-original-html");
  if (original != null) {
    tbody.innerHTML = original;
    tbody.removeAttribute("data-original-html");
  }
  const ph = document.getElementById("printHeader");
  if (ph) ph.remove();
}

/* =========================
   Events
========================= */
const onChange = debounce(render, 160);
els.qName.addEventListener("input", onChange);
els.qId.addEventListener("input", onChange);
els.qStatus.addEventListener("change", render);
els.qRole.addEventListener("change", render);

// group dropdown open/close
els.btnGroupDD.addEventListener("click", () => openGroupDD());
document.addEventListener("click", (e) => {
  if (!els.groupDD.contains(e.target)) openGroupDD(false);
});

// group search + checkboxes
els.groupSearch.addEventListener("input", () => renderGroupList());

els.groupList.addEventListener("change", (e) => {
  const cb = e.target.closest('input[type="checkbox"][data-group]');
  if (!cb) return;
  const key = norm(cb.getAttribute("data-group"));
  if (cb.checked) GROUP_SELECTED.add(key);
  else GROUP_SELECTED.delete(key);
  updateGroupDDText();
  render();
});

els.btnGroupAll.addEventListener("click", () => setAllGroupsSelected(true));
els.btnGroupNone.addEventListener("click", () => setAllGroupsSelected(false));

els.btnClear.addEventListener("click", () => {
  els.qName.value = "";
  els.qId.value = "";
  els.qStatus.value = "";
  els.qRole.value = "";
  GROUP_SELECTED.clear();
  updateGroupDDText();
  renderGroupList();
  render();
});

els.btnReload.addEventListener("click", async () => {
  try {
    await fetchSheet();
  } catch (e) {
    console.error(e);
    setStatus("Load failed (check sheet sharing/public)", "err");
    els.tbody.innerHTML = `
      <tr class="empty"><td colspan="9">
        Load បរាជ័យ។ សូមធ្វើឲ្យ Sheet អាចអានបាន (Anyone with link / Published) ហើយ Reload ម្តងទៀត
      </td></tr>`;
  }
});

els.btnPrint.addEventListener("click", () => {
  const rowsToPrint = hasActiveFilter() ? (LAST_FILTERED || []) : (CACHE.rows || []);
  preparePrint(rowsToPrint);

  window.onafterprint = () => {
    cleanupPrint();
    window.onafterprint = null;
  };

  window.print();
});

// Auto load
fetchSheet().catch(() => setStatus("Ready", "idle"));
