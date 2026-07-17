// ============================================================
// app.js — auth, data load, filters, all dashboard views,
// drill-down, explorer, export.
// ============================================================
/* global CONFIG, Chart, XLSX, supabase, Upload */

const S = {
  sb: null, session: null, role: "viewer",
  tickets: [], storeCount: 0, asOn: null,
  charts: {}, agg: {},          // agg tables cached for CSV export
  page: 0, pageSize: 50,
  sort: {},                     // per-table sort state
  activeTab: "overview", dirty: true,
};

const FILTER_DEFS = [
  ["f-year",   "year",   t => t.issue_raised_year],
  ["f-quarter","quarter",t => t.quarter_raised],
  ["f-region", "region", t => t.region],
  ["f-branch", "branch", t => t.branch],
  ["f-tier",   "tier",   t => t.city_classification],
  ["f-budget", "budget", t => t.budget_category],
  ["f-status", "status", t => t.status],
  ["f-final",  "final",  t => t.final_status],
  ["f-resp",   "resp",   t => t.responsibility],
  ["f-logo",   "logo",   t => t.logo_flag],
];
const F = { year:"", quarter:"", region:"", branch:"", tier:"", budget:"", status:"", final:"", resp:"", logo:"", search:"" };

const $ = id => document.getElementById(id);
const fmt = n => n == null ? "-" : Number(n).toLocaleString("en-IN");
const fmt1 = n => n == null || isNaN(n) ? "-" : Number(n).toFixed(1);
const pct = (a, b) => b ? (100 * a / b).toFixed(1) + "%" : "-";
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));

// ---------------- boot ----------------
window.addEventListener("DOMContentLoaded", async () => {
  if (String(CONFIG.SUPABASE_URL).includes("PASTE_")) {
    document.body.innerHTML = '<div style="max-width:560px;margin:80px auto;font-family:sans-serif;line-height:1.6"><h2>⚙ Setup needed</h2><p>Open <code>js/config.js</code> and paste your Supabase project URL and anon key, then run <code>supabase/schema.sql</code> in the Supabase SQL editor. See README.md for the full guide.</p></div>';
    return;
  }
  S.sb = supabase.createClient(CONFIG.SUPABASE_URL, CONFIG.SUPABASE_ANON_KEY);
  initTheme(); bindUI();
  const { data: { session } } = await S.sb.auth.getSession();
  session ? enterApp(session) : showLogin();
  S.sb.auth.onAuthStateChange((_e, sess) => {
    if (sess && !S.session) enterApp(sess);
    if (!sess) { S.session = null; showLogin(); }
  });
});

function showLogin() { $("login-screen").classList.remove("hidden"); $("app").classList.add("hidden"); }

async function enterApp(session) {
  S.session = session;
  $("login-screen").classList.add("hidden");
  $("app").classList.remove("hidden");
  $("user-email").textContent = session.user.email;
  const { data: prof } = await S.sb.from("profiles").select("role").eq("id", session.user.id).single();
  S.role = prof ? prof.role : "viewer";
  if (S.role === "uploader" || S.role === "admin") $("upload-tab-btn").classList.remove("hidden");
  await loadData();
}

// ---------------- data load ----------------
async function loadData() {
  loading(true, "Loading tickets…");
  try {
    const all = []; const CHUNK = 1000;
    for (let from = 0; ; from += CHUNK) {
      const { data, error } = await S.sb.from("tickets").select("*").order("ticket_id").range(from, from + CHUNK - 1);
      if (error) throw error;
      all.push(...data);
      loading(true, `Loading tickets… ${all.length}`);
      if (data.length < CHUNK) break;
    }
    S.tickets = all;
    const [{ count: storeCount }, { data: lastUp }] = await Promise.all([
      S.sb.from("stores").select("*", { count: "exact", head: true }),
      S.sb.from("upload_logs").select("as_on_date,uploaded_at").order("uploaded_at", { ascending: false }).limit(1),
    ]);
    S.storeCount = storeCount || 0;
    S.asOn = (lastUp && lastUp[0] && lastUp[0].as_on_date) || null;
    $("as-on-badge").textContent = S.asOn ? "Data as on " + fmtDate(S.asOn) : (all.length ? "" : "No data yet — use Upload tab");
    buildFilterOptions();
    S.dirty = true; renderActive();
    if (S.role !== "viewer") Upload.loadHistory();
  } catch (e) {
    toast("Failed to load data: " + e.message, 5000);
  } finally { loading(false); }
}

function fmtDate(d) { if (!d) return "-"; const [y, m, dd] = d.split("-"); return `${dd}-${m}-${y}`; }

// ---------------- filters ----------------
function buildFilterOptions() {
  for (const [elId, , getter] of FILTER_DEFS) {
    const sel = $(elId), cur = sel.value;
    const vals = [...new Set(S.tickets.map(getter).filter(v => v != null && v !== ""))];
    vals.sort((a, b) => elId === "f-quarter" ? qKey(a) - qKey(b) : String(a).localeCompare(String(b)));
    sel.innerHTML = '<option value="">All</option>' + vals.map(v => `<option>${esc(v)}</option>`).join("");
    if (vals.includes(isNaN(cur) ? cur : Number(cur)) || vals.includes(cur)) sel.value = cur;
  }
}

function qKey(q) { const m = String(q).match(/Q(\d)\s*(\d{4})/); return m ? Number(m[2]) * 4 + Number(m[1]) : 99999; }

function readFilters() {
  for (const [elId, key] of FILTER_DEFS) F[key] = $(elId).value;
  F.search = $("f-search").value.trim().toLowerCase();
  renderChips();
  S.page = 0; S.dirty = true; renderActive();
}

function setFilters(patch) {  // used by drill-down
  Object.keys(F).forEach(k => { if (k in patch) F[k] = patch[k] == null ? "" : String(patch[k]); });
  for (const [elId, key] of FILTER_DEFS) if (key in patch) $(elId).value = F[key];
  if ("search" in patch) $("f-search").value = F.search;
  renderChips(); S.page = 0; S.dirty = true;
}

function clearFilters() {
  Object.keys(F).forEach(k => F[k] = "");
  for (const [elId] of FILTER_DEFS) $(elId).value = "";
  $("f-search").value = "";
  renderChips(); S.page = 0; S.dirty = true; renderActive();
}

function renderChips() {
  const box = $("active-filters"); box.innerHTML = "";
  const labels = { year:"Year", quarter:"Quarter", region:"Region", branch:"Branch", tier:"Tier", budget:"Budget", status:"Stage", final:"Status", resp:"Resp.", logo:"Logo", search:"Search" };
  Object.entries(F).forEach(([k, v]) => {
    if (!v) return;
    const chip = document.createElement("span");
    chip.className = "chip"; chip.textContent = `${labels[k]}: ${v} ✕`;
    chip.onclick = () => { setFilters({ [k]: "" }); renderActive(); };
    box.appendChild(chip);
  });
}

function getFiltered() {
  return S.tickets.filter(t =>
    (!F.year    || String(t.issue_raised_year) === F.year) &&
    (!F.quarter || t.quarter_raised === F.quarter) &&
    (!F.region  || t.region === F.region) &&
    (!F.branch  || t.branch === F.branch) &&
    (!F.tier    || t.city_classification === F.tier) &&
    (!F.budget  || t.budget_category === F.budget) &&
    (!F.status  || t.status === F.status) &&
    (!F.final   || t.final_status === F.final) &&
    (!F.resp    || t.responsibility === F.resp) &&
    (!F.logo    || t.logo_flag === F.logo) &&
    (!F.search  || matchesSearch(t, F.search))
  );
}
function matchesSearch(t, q) {
  return [t.ticket_id, t.new_ticket_no, t.store_name, t.city, t.branch, t.state, t.issue_category, t.problem_reported]
    .some(v => v != null && String(v).toLowerCase().includes(q));
}

// ---------------- UI bindings ----------------
function bindUI() {
  $("login-form").addEventListener("submit", async e => {
    e.preventDefault();
    $("login-btn").disabled = true; $("login-error").textContent = "";
    const { error } = await S.sb.auth.signInWithPassword({ email: $("login-email").value, password: $("login-password").value });
    $("login-btn").disabled = false;
    if (error) $("login-error").textContent = error.message;
  });
  $("signout-btn").onclick = () => S.sb.auth.signOut();
  $("theme-toggle").onclick = toggleTheme;

  document.querySelectorAll("#tabs .tab").forEach(b => b.onclick = () => switchTab(b.dataset.tab));
  for (const [elId] of FILTER_DEFS) $(elId).addEventListener("change", readFilters);
  $("f-search").addEventListener("input", debounce(readFilters, 300));
  $("filters-clear").onclick = clearFilters;

  // presets
  $("preset-save").onclick = () => {
    const name = prompt("Preset name:"); if (!name) return;
    const all = JSON.parse(localStorage.getItem("asus_presets") || "{}");
    all[name] = { ...F }; localStorage.setItem("asus_presets", JSON.stringify(all));
    loadPresetList(); toast("Preset saved");
  };
  $("preset-list").onchange = e => {
    const all = JSON.parse(localStorage.getItem("asus_presets") || "{}");
    if (all[e.target.value]) { setFilters(all[e.target.value]); renderActive(); }
    e.target.value = "";
  };
  loadPresetList();

  $("pg-prev").onclick = () => { if (S.page > 0) { S.page--; renderExplorer(); } };
  $("pg-next").onclick = () => { S.page++; renderExplorer(); };
  $("export-xlsx").onclick = () => exportRaw("xlsx");
  $("export-csv").onclick = () => exportRaw("csv");
  document.querySelectorAll(".export-agg").forEach(b => b.onclick = ev => { ev.stopPropagation(); exportAgg(b.dataset.agg); });

  $("store-search").addEventListener("input", debounce(() => renderStores(), 250));
  $("city-search").addEventListener("input", debounce(() => renderStores(), 250));

  $("drawer-close").onclick = closeDrawer;
  $("drawer-overlay").onclick = closeDrawer;

  Upload.bind();
}

function loadPresetList() {
  const all = JSON.parse(localStorage.getItem("asus_presets") || "{}");
  $("preset-list").innerHTML = '<option value="">Presets…</option>' + Object.keys(all).map(n => `<option>${esc(n)}</option>`).join("");
}

function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }

function switchTab(tab) {
  S.activeTab = tab;
  document.querySelectorAll("#tabs .tab").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll(".view").forEach(v => v.classList.add("hidden"));
  $("view-" + tab).classList.remove("hidden");
  renderActive(true);
}

function renderActive(force) {
  if (!S.dirty && !force) return;
  const t = S.activeTab;
  if (t === "overview") renderOverview();
  else if (t === "regional") renderRegional();
  else if (t === "stores") renderStores();
  else if (t === "issues") renderIssues();
  else if (t === "ageing") renderAgeing();
  else if (t === "explorer") renderExplorer();
  if (t !== "upload") S.dirty = false;
}

// ---------------- aggregation helpers ----------------
function groupBy(rows, keyFn) {
  const m = new Map();
  for (const r of rows) { const k = keyFn(r) ?? "(blank)"; (m.get(k) || m.set(k, []).get(k)).push(r); }
  return m;
}
function avg(rows, f) { const v = rows.map(f).filter(x => x != null && !isNaN(x)); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; }
function stat(rows) {
  const closed = rows.filter(r => r.final_status === "Closed");
  const open = rows.filter(r => r.final_status === "Open");
  const inTat = closed.filter(r => r.tat_follow === "InTAT").length;
  return { total: rows.length, closed: closed.length, open: open.length,
    avgTat: avg(closed, r => r.rectification_time), inTat,
    stores: new Set(rows.map(r => r.store_name).filter(Boolean)).size };
}
const openAgeDays = t => {
  const ref = S.asOn ? new Date(S.asOn) : new Date();
  return t.issue_raised_date ? Math.max(0, Math.round((ref - new Date(t.issue_raised_date)) / 86400000)) : null;
};
const OPEN_BUCKETS = [["00 to 10 Days",0,10],["11 to 20 Days",11,20],["21 to 30 Days",21,30],["31 to 40 Days",31,40],["41 to 50 Days",41,50],["51 to 60 Days",51,60],["61 to 90 Days",61,90],["91 to 150 Days",91,150],["Above 150 Days",151,1e9]];
const openBucket = d => { if (d == null) return "(no date)"; const b = OPEN_BUCKETS.find(([, lo, hi]) => d >= lo && d <= hi); return b ? b[0] : "(no date)"; };

// palette
const PALETTE = ["#0057d8","#00a29a","#f2a53a","#d4380d","#7b61c9","#1f9d55","#e05d8a","#5b8dec","#98a1b3","#c4b13d"];
function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }

function makeChart(id, cfg, onLabelClick) {
  if (S.charts[id]) S.charts[id].destroy();
  if (typeof ChartDataLabels !== "undefined" && !S._dlReg) { Chart.register(ChartDataLabels); S._dlReg = true; }
  const base = {
    responsive: true, maintainAspectRatio: true,
    plugins: { legend: { labels: { color: cssVar("--text") } } },
  };
  cfg.options = deepMerge(base, cfg.options || {});
  const isPie = cfg.type === "doughnut" || cfg.type === "pie";
  const horizontal = cfg.options.indexAxis === "y";
  const stacked = !!(cfg.options.scales && ((cfg.options.scales.x || {}).stacked || (cfg.options.scales.y || {}).stacked));
  if (!isPie) {
    cfg.options.scales = deepMerge({
      x: { ticks: { color: cssVar("--muted") }, grid: { color: cssVar("--border") } },
      y: { ticks: { color: cssVar("--muted") }, grid: { color: cssVar("--border") } },
    }, cfg.options.scales || {});
    // headroom so value labels above bars/points don't get clipped
    cfg.options.scales[horizontal ? "x" : "y"].grace = stacked ? undefined : "10%";
  }
  // show numbers on bars / points / slices (skipped gracefully if plugin CDN failed)
  cfg.options.plugins.datalabels = deepMerge({
    display: c => { const v = c.dataset.data[c.dataIndex]; return v != null && v !== 0; },
    color: isPie ? "#fff" : cssVar("--text"),
    font: { size: 10.5, weight: "600" },
    formatter: v => typeof v === "number" ? (Number.isInteger(v) ? v.toLocaleString("en-IN") : v.toFixed(1)) : v,
    anchor: (isPie || stacked) ? "center" : "end",
    align: (isPie || stacked) ? "center" : "end",
    offset: (isPie || stacked) ? 0 : 2,
    clamp: true, clip: false,
  }, cfg.options.plugins.datalabels || {});
  if (onLabelClick) {
    cfg.options.onClick = (evt, els, chart) => {
      const pts = chart.getElementsAtEventForMode(evt, "nearest", { intersect: true }, true);
      if (pts.length) onLabelClick(chart.data.labels[pts[0].index], pts[0].datasetIndex);
    };
    cfg.options.onHover = (e, els) => { e.native.target.style.cursor = els.length ? "pointer" : "default"; };
  }
  S.charts[id] = new Chart($(id), cfg);
}
function deepMerge(a, b) {
  const out = { ...a };
  for (const k of Object.keys(b)) out[k] = (b[k] && typeof b[k] === "object" && !Array.isArray(b[k]) && a[k]) ? deepMerge(a[k], b[k]) : b[k];
  return out;
}

function drill(patch, label) {
  setFilters(patch);
  switchTab("explorer");
  toast("Drilled into: " + (label || "selection") + " — showing underlying tickets");
}

// ============================================================
// OVERVIEW
// ============================================================
function renderOverview() {
  const rows = getFiltered();
  const s = stat(rows);
  const cp = rows.filter(r => r.responsibility === "Channelplay");
  const rv = rows.filter(r => r.responsibility && r.responsibility !== "Channelplay");
  const cpClosed = cp.filter(r => r.final_status === "Closed");
  const rvClosed = rv.filter(r => r.final_status === "Closed");

  const kpis = [
    { label: "Total Tickets", value: fmt(s.total), drill: {} },
    { label: "Closed", value: fmt(s.closed), cls: "k-ok", sub: pct(s.closed, s.total) + " closure", drill: { final: "Closed" } },
    { label: "Open", value: fmt(s.open), cls: "k-bad", drill: { final: "Open" } },
    { label: "Avg TAT (days)", value: fmt1(s.avgTat), sub: "closed tickets" },
    { label: "In-TAT Closed", value: fmt(s.inTat), cls: "k-ok", sub: pct(s.inTat, s.closed) + " of closed" },
    { label: "CP Avg TAT", value: fmt1(avg(cpClosed, r => r.rectification_time)), sub: fmt(cpClosed.length) + " closed by CP", drill: { resp: "Channelplay" } },
    { label: "ASUS/RV Avg TAT", value: fmt1(avg(rvClosed, r => r.rectification_time)), sub: fmt(rvClosed.length) + " closed by ASUS/RV", drill: { resp: "Asus" } },
    { label: "Stores w/ Tickets", value: fmt(s.stores), sub: S.storeCount ? "of " + fmt(S.storeCount) + " stores (" + pct(s.stores, S.storeCount) + ")" : "" },
  ];
  $("kpi-grid").innerHTML = kpis.map((k, i) =>
    `<div class="kpi ${k.cls || ""}" data-i="${i}"><div class="kpi-label">${k.label}</div><div class="kpi-value">${k.value}</div><div class="kpi-sub">${k.sub || ""}</div></div>`).join("");
  $("kpi-grid").querySelectorAll(".kpi").forEach((el, i) => { if (kpis[i].drill) el.onclick = () => drill(kpis[i].drill, kpis[i].label); });

  // quarterly chart
  const quarters = [...new Set(rows.map(r => r.quarter_raised).filter(Boolean))].sort((a, b) => qKey(a) - qKey(b));
  const recvByQ = quarters.map(q => rows.filter(r => r.quarter_raised === q).length);
  const closByQ = quarters.map(q => rows.filter(r => r.quarter_rectified === q && r.final_status === "Closed").length);
  const openByQ = quarters.map(q => rows.filter(r => r.quarter_raised === q && r.final_status === "Open").length);
  S.agg.quarterly = { headers: ["Quarter","Received","Closed in Qtr","Still Open (raised in Qtr)"], rows: quarters.map((q, i) => [q, recvByQ[i], closByQ[i], openByQ[i]]) };
  makeChart("ch-quarterly", { type: "bar",
    data: { labels: quarters, datasets: [
      { label: "Received", data: recvByQ, backgroundColor: PALETTE[0] },
      { label: "Closed (in qtr)", data: closByQ, backgroundColor: PALETTE[5] },
      { label: "Still open", data: openByQ, backgroundColor: PALETTE[3] } ] } },
    (label, dsi) => drill(dsi === 2 ? { quarter: label, final: "Open" } : { quarter: label }, label));

  // stage funnel
  const stages = [...groupBy(rows, r => r.status).entries()].sort((a, b) => b[1].length - a[1].length);
  S.agg.stages = { headers: ["Stage","Tickets"], rows: stages.map(([k, v]) => [k, v.length]) };
  makeChart("ch-stages", { type: "bar",
    data: { labels: stages.map(s2 => s2[0]), datasets: [{ label: "Tickets", data: stages.map(s2 => s2[1].length), backgroundColor: PALETTE }] },
    options: { indexAxis: "y", plugins: { legend: { display: false } } } },
    label => drill({ status: label }, label));

  // budget donut
  const bud = [...groupBy(rows, r => r.budget_category).entries()].sort((a, b) => b[1].length - a[1].length);
  S.agg.budget = { headers: ["Budget Category","Tickets","Share %"], rows: bud.map(([k, v]) => [k, v.length, (100 * v.length / (s.total || 1)).toFixed(1)]) };
  makeChart("ch-budget", { type: "doughnut",
    data: { labels: bud.map(b => b[0]), datasets: [{ data: bud.map(b => b[1].length), backgroundColor: PALETTE }] } },
    label => drill({ budget: label }, label));

  // monthly trend
  const mKey = d => d ? d.slice(0, 7) : null;
  const months = [...new Set([...rows.map(r => mKey(r.issue_raised_date)), ...rows.map(r => mKey(r.rectification_date))].filter(Boolean))].sort();
  const raised = months.map(m => rows.filter(r => mKey(r.issue_raised_date) === m).length);
  const rect = months.map(m => rows.filter(r => mKey(r.rectification_date) === m && r.final_status === "Closed").length);
  S.agg.monthly = { headers: ["Month","Raised","Rectified"], rows: months.map((m, i) => [m, raised[i], rect[i]]) };
  makeChart("ch-monthly", { type: "line",
    data: { labels: months, datasets: [
      { label: "Raised", data: raised, borderColor: PALETTE[0], backgroundColor: PALETTE[0], tension: .3 },
      { label: "Rectified", data: rect, borderColor: PALETTE[5], backgroundColor: PALETTE[5], tension: .3 } ] } });

  // quarterly detail table
  const qd = quarters.map(q => {
    const recv = rows.filter(r => r.quarter_raised === q);
    const closedInQ = rows.filter(r => r.quarter_rectified === q && r.final_status === "Closed");
    const cpQ = closedInQ.filter(r => r.responsibility === "Channelplay");
    const rvQ = closedInQ.filter(r => r.responsibility !== "Channelplay");
    return { q, recv: recv.length, closed: closedInQ.length,
      cp: cpQ.length, cpTat: avg(cpQ, r => r.rectification_time),
      rv: rvQ.length, rvTat: avg(rvQ, r => r.rectification_time),
      inTat: closedInQ.filter(r => r.tat_follow === "InTAT").length,
      open: recv.filter(r => r.final_status === "Open").length };
  });
  S.agg.quarterlyDetail = { headers: ["Quarter","Received","Closed in Qtr","Closed by CP","CP Avg TAT","Closed by ASUS/RV","RV Avg TAT","In-TAT Closed","Still Open"],
    rows: qd.map(r => [r.q, r.recv, r.closed, r.cp, fmt1(r.cpTat), r.rv, fmt1(r.rvTat), r.inTat, r.open]) };
  renderTable("tbl-quarterly", S.agg.quarterlyDetail.headers, qd.map(r => ({
    cells: [r.q, fmt(r.recv), fmt(r.closed), fmt(r.cp), fmt1(r.cpTat), fmt(r.rv), fmt1(r.rvTat), fmt(r.inTat), `<span class="pill open">${fmt(r.open)}</span>`],
    numCols: [1,2,3,4,5,6,7,8], onClick: () => drill({ quarter: r.q }, r.q) })),
    { totals: ["Total", fmt(qd.reduce((a, r) => a + r.recv, 0)), fmt(qd.reduce((a, r) => a + r.closed, 0)), fmt(qd.reduce((a, r) => a + r.cp, 0)), "", fmt(qd.reduce((a, r) => a + r.rv, 0)), "", fmt(qd.reduce((a, r) => a + r.inTat, 0)), fmt(qd.reduce((a, r) => a + r.open, 0))] });

  // half-yearly
  const hyOf = d => d ? d.slice(0, 4) + (Number(d.slice(5, 7)) <= 6 ? " HY1" : " HY2") : null;
  const hys = [...new Set([...rows.map(r => hyOf(r.issue_raised_date)), ...rows.map(r => r.final_status === "Closed" ? hyOf(r.rectification_date) : null)].filter(Boolean))].sort();
  const hyRows = hys.map(h => {
    const recv = rows.filter(r => hyOf(r.issue_raised_date) === h);
    const clos = rows.filter(r => r.final_status === "Closed" && hyOf(r.rectification_date) === h);
    return { h, recv: recv.length, closed: clos.length, tat: avg(clos, r => r.rectification_time) };
  });
  S.agg.halfyear = { headers: ["Half Year","Received","Closed","Avg TAT (days)"], rows: hyRows.map(r => [r.h, r.recv, r.closed, fmt1(r.tat)]) };
  renderTable("tbl-halfyear", S.agg.halfyear.headers, hyRows.map((r, i) => {
    const prev = hyRows[i - 1];
    const delta = prev ? ((r.recv - prev.recv) >= 0 ? " ▲" : " ▼") : "";
    return { cells: [r.h, fmt(r.recv) + delta, fmt(r.closed), fmt1(r.tat)], numCols: [1,2,3] };
  }));
}

// ============================================================
// REGIONAL
// ============================================================
function renderRegional() {
  const rows = getFiltered();
  const regs = [...groupBy(rows, r => r.region).entries()].sort((a, b) => b[1].length - a[1].length);
  makeChart("ch-region", { type: "bar",
    data: { labels: regs.map(r => r[0]), datasets: [
      { label: "Open", data: regs.map(([, v]) => v.filter(t => t.final_status === "Open").length), backgroundColor: PALETTE[3] },
      { label: "Closed", data: regs.map(([, v]) => v.filter(t => t.final_status === "Closed").length), backgroundColor: PALETTE[5] } ] },
    options: { scales: { x: { stacked: true }, y: { stacked: true } } } },
    (label, dsi) => drill({ region: label, final: dsi === 0 ? "Open" : "Closed" }, label));

  makeChart("ch-region-tat", { type: "bar",
    data: { labels: regs.map(r => r[0]), datasets: [{ label: "Avg TAT (days)", data: regs.map(([, v]) => avg(v.filter(t => t.final_status === "Closed"), t => t.rectification_time)), backgroundColor: PALETTE[1] }] },
    options: { plugins: { legend: { display: false } } } },
    label => drill({ region: label, final: "Closed" }, label));

  const regRows = regs.map(([k, v]) => { const st = stat(v); return { k, ...st }; });
  S.agg.region = { headers: ["Region","Received","Closed","Open","Closure %","Avg TAT","In-TAT % (closed)","Stores Affected"],
    rows: regRows.map(r => [r.k, r.total, r.closed, r.open, pct(r.closed, r.total), fmt1(r.avgTat), pct(r.inTat, r.closed), r.stores]) };
  renderTable("tbl-region", S.agg.region.headers, regRows.map(r => ({
    cells: [r.k, fmt(r.total), fmt(r.closed), `<span class="pill open">${fmt(r.open)}</span>`, pct(r.closed, r.total), fmt1(r.avgTat), pct(r.inTat, r.closed), fmt(r.stores)],
    numCols: [1,2,3,4,5,6,7], onClick: () => drill({ region: r.k }, r.k) })));

  const brs = [...groupBy(rows, r => r.branch).entries()].sort((a, b) => b[1].length - a[1].length);
  const brRows = brs.map(([k, v]) => ({ k, region: v[0].region, ...stat(v) }));
  S.agg.branch = { headers: ["Branch","Region","Received","Closed","Open","Closure %","Avg TAT","Stores Affected"],
    rows: brRows.map(r => [r.k, r.region, r.total, r.closed, r.open, pct(r.closed, r.total), fmt1(r.avgTat), r.stores]) };
  renderTable("tbl-branch", S.agg.branch.headers, brRows.map(r => ({
    cells: [r.k, r.region || "-", fmt(r.total), fmt(r.closed), `<span class="pill open">${fmt(r.open)}</span>`, pct(r.closed, r.total), fmt1(r.avgTat), fmt(r.stores)],
    numCols: [2,3,4,5,6,7], onClick: () => drill({ branch: r.k }, r.k) })));
}

// ============================================================
// STORES & CITIES
// ============================================================
function renderStores() {
  const rows = getFiltered();
  const stores = [...groupBy(rows, r => r.store_name).entries()].sort((a, b) => b[1].length - a[1].length);
  const cities = [...groupBy(rows, r => r.city).entries()].sort((a, b) => b[1].length - a[1].length);

  $("store-kpis").innerHTML = [
    ["Store Universe", S.storeCount ? fmt(S.storeCount) : "—", "from store master"],
    ["Stores with Tickets", fmt(stores.length), S.storeCount ? pct(stores.length, S.storeCount) + " of universe" : ""],
    ["Stores w/ Open Tickets", fmt(new Set(rows.filter(r => r.final_status === "Open").map(r => r.store_name)).size), ""],
    ["Cities Covered", fmt(cities.length), ""],
    ["Repeat-Issue Stores", fmt(stores.filter(([, v]) => v.length >= 3).length), "3+ tickets"],
  ].map(([l, v, sub]) => `<div class="kpi"><div class="kpi-label">${l}</div><div class="kpi-value">${v}</div><div class="kpi-sub">${sub}</div></div>`).join("");

  const top = stores.slice(0, 15);
  makeChart("ch-topstores", { type: "bar",
    data: { labels: top.map(s => s[0]), datasets: [{ label: "Tickets", data: top.map(s => s[1].length), backgroundColor: PALETTE[0] }] },
    options: { indexAxis: "y", plugins: { legend: { display: false } } } },
    label => drill({ search: label.toLowerCase() }, label));

  const tiers = [...groupBy(rows, r => r.city_classification).entries()];
  makeChart("ch-tier", { type: "doughnut",
    data: { labels: tiers.map(t => t[0]), datasets: [{ data: tiers.map(t => t[1].length), backgroundColor: PALETTE }] } },
    label => drill({ tier: label }, label));

  const sq = $("store-search").value.trim().toLowerCase();
  const stRows = stores.filter(([k]) => !sq || String(k).toLowerCase().includes(sq)).slice(0, 400)
    .map(([k, v]) => ({ k, city: v[0].city, region: v[0].region, branch: v[0].branch, type: v[0].store_type, ...stat(v) }));
  S.agg.store = { headers: ["Store","City","Region","Branch","Type","Tickets","Open","Closed","Avg TAT"],
    rows: stRows.map(r => [r.k, r.city, r.region, r.branch, r.type, r.total, r.open, r.closed, fmt1(r.avgTat)]) };
  renderTable("tbl-store", S.agg.store.headers, stRows.map(r => ({
    cells: [r.k, r.city || "-", r.region || "-", r.branch || "-", r.type || "-", fmt(r.total), `<span class="pill open">${fmt(r.open)}</span>`, fmt(r.closed), fmt1(r.avgTat)],
    numCols: [5,6,7,8], onClick: () => drill({ search: String(r.k).toLowerCase() }, r.k) })));

  const cq = $("city-search").value.trim().toLowerCase();
  const ctRows = cities.filter(([k]) => !cq || String(k).toLowerCase().includes(cq)).slice(0, 400)
    .map(([k, v]) => ({ k, tier: v[0].city_classification, state: v[0].state, region: v[0].region, ...stat(v) }));
  S.agg.city = { headers: ["City","Tier","State","Region","Tickets","Open","Closed","Avg TAT","Stores"],
    rows: ctRows.map(r => [r.k, r.tier, r.state, r.region, r.total, r.open, r.closed, fmt1(r.avgTat), r.stores]) };
  renderTable("tbl-city", S.agg.city.headers, ctRows.map(r => ({
    cells: [r.k, r.tier || "-", r.state || "-", r.region || "-", fmt(r.total), `<span class="pill open">${fmt(r.open)}</span>`, fmt(r.closed), fmt1(r.avgTat), fmt(r.stores)],
    numCols: [4,5,6,7,8], onClick: () => drill({ search: String(r.k).toLowerCase() }, r.k) })));
}

// ============================================================
// ISSUES & BOTTLENECKS
// ============================================================
function renderIssues() {
  const rows = getFiltered();

  // budget category — received vs closed vs open
  const buds = [...groupBy(rows, r => r.budget_category).entries()].sort((a, b) => b[1].length - a[1].length);
  S.agg.budgetStatus = { headers: ["Budget Category","Received","Closed","Open"],
    rows: buds.map(([k, v]) => [k, v.length, v.filter(t => t.final_status === "Closed").length, v.filter(t => t.final_status === "Open").length]) };
  makeChart("ch-budget-status", { type: "bar",
    data: { labels: buds.map(b => b[0]), datasets: [
      { label: "Received", data: buds.map(([, v]) => v.length), backgroundColor: PALETTE[0] },
      { label: "Closed", data: buds.map(([, v]) => v.filter(t => t.final_status === "Closed").length), backgroundColor: PALETTE[5] },
      { label: "Open", data: buds.map(([, v]) => v.filter(t => t.final_status === "Open").length), backgroundColor: PALETTE[3] } ] } },
    (label, dsi) => drill(dsi === 0 ? { budget: label } : { budget: label, final: dsi === 1 ? "Closed" : "Open" }, label));

  const cats = [...groupBy(rows, r => r.issue_category).entries()].sort((a, b) => b[1].length - a[1].length);
  const top12 = cats.slice(0, 12);
  makeChart("ch-issuecat", { type: "bar",
    data: { labels: top12.map(c => c[0]), datasets: [{ label: "Tickets", data: top12.map(c => c[1].length), backgroundColor: PALETTE[0] }] },
    options: { indexAxis: "y", plugins: { legend: { display: false } } } },
    label => drill({ search: label.toLowerCase() }, label));

  const openCats = [...groupBy(rows.filter(r => r.final_status === "Open"), r => r.issue_category).entries()].sort((a, b) => b[1].length - a[1].length).slice(0, 12);
  makeChart("ch-opencat", { type: "bar",
    data: { labels: openCats.map(c => c[0]), datasets: [{ label: "Open tickets", data: openCats.map(c => c[1].length), backgroundColor: PALETTE[3] }] },
    options: { indexAxis: "y", plugins: { legend: { display: false } } } },
    label => drill({ search: label.toLowerCase(), final: "Open" }, label));

  const slow = cats.map(([k, v]) => [k, avg(v.filter(t => t.final_status === "Closed"), t => t.rectification_time), v.length])
    .filter(x => x[1] != null && x[2] >= 10).sort((a, b) => b[1] - a[1]).slice(0, 12);
  makeChart("ch-cat-tat", { type: "bar",
    data: { labels: slow.map(s => s[0]), datasets: [{ label: "Avg days to close", data: slow.map(s => s[1]), backgroundColor: PALETTE[2] }] },
    options: { indexAxis: "y", plugins: { legend: { display: false } } } },
    label => drill({ search: label.toLowerCase(), final: "Closed" }, label));

  // stuck: open by stage × ageing bucket
  const open = rows.filter(r => r.final_status === "Open");
  const stages = [...groupBy(open, r => r.status).entries()].sort((a, b) => b[1].length - a[1].length);
  const bucketNames = OPEN_BUCKETS.map(b => b[0]);
  const stuckRows = stages.map(([st, v]) => {
    const counts = bucketNames.map(bn => v.filter(t => openBucket(openAgeDays(t)) === bn).length);
    return { st, total: v.length, counts };
  });
  S.agg.stuck = { headers: ["Stage", "Open", ...bucketNames], rows: stuckRows.map(r => [r.st, r.total, ...r.counts]) };
  renderTable("tbl-stuck", ["Stage", "Open", ...bucketNames.map(b => b.replace(" Days",""))], stuckRows.map(r => ({
    cells: [r.st, fmt(r.total), ...r.counts.map(fmt)], numCols: Array.from({length: bucketNames.length + 1}, (_, i) => i + 1),
    onClick: () => drill({ status: r.st, final: "Open" }, r.st) })));

  const catRows = cats.map(([k, v]) => ({ k, budget: v[0].budget_category, ...stat(v) }));
  S.agg.issuecat = { headers: ["Issue Category","Budget Cat.","Tickets","Share %","Open","Closed","Avg TAT"],
    rows: catRows.map(r => [r.k, r.budget, r.total, pct(r.total, rows.length), r.open, r.closed, fmt1(r.avgTat)]) };
  renderTable("tbl-issuecat", S.agg.issuecat.headers, catRows.map(r => ({
    cells: [r.k, r.budget || "-", fmt(r.total), pct(r.total, rows.length), `<span class="pill open">${fmt(r.open)}</span>`, fmt(r.closed), fmt1(r.avgTat)],
    numCols: [2,3,4,5,6], onClick: () => drill({ search: String(r.k).toLowerCase() }, r.k) })));
}

// ============================================================
// AGEING & TAT
// ============================================================
function renderAgeing() {
  const rows = getFiltered();
  const open = rows.filter(r => r.final_status === "Open");
  const bucketNames = OPEN_BUCKETS.map(b => b[0]);
  const openCounts = bucketNames.map(bn => open.filter(t => openBucket(openAgeDays(t)) === bn).length);
  makeChart("ch-open-ageing", { type: "bar",
    data: { labels: bucketNames, datasets: [{ label: "Open tickets", data: openCounts, backgroundColor: PALETTE[3] }] },
    options: { plugins: { legend: { display: false } } } },
    label => { setFilters({ final: "Open" }); switchTab("explorer"); toast("Open tickets — check Ageing column for " + label); });

  const closed = rows.filter(r => r.final_status === "Closed");
  const cBuckets = [...new Set(closed.map(r => r.ageing_closure_bucket).filter(Boolean))].sort();
  makeChart("ch-close-ageing", { type: "bar",
    data: { labels: cBuckets, datasets: [
      { label: "Channelplay", data: cBuckets.map(b => closed.filter(r => r.ageing_closure_bucket === b && r.responsibility === "Channelplay").length), backgroundColor: PALETTE[0] },
      { label: "ASUS/RV", data: cBuckets.map(b => closed.filter(r => r.ageing_closure_bucket === b && r.responsibility !== "Channelplay").length), backgroundColor: PALETTE[1] } ] },
    options: { scales: { x: { stacked: true }, y: { stacked: true } } } },
    label => drill({ final: "Closed" }, label));

  const quarters = [...new Set(closed.map(r => r.quarter_rectified).filter(Boolean))].sort((a, b) => qKey(a) - qKey(b));
  makeChart("ch-tat-comp", { type: "bar",
    data: { labels: quarters, datasets: [
      { label: "In TAT", data: quarters.map(q => closed.filter(r => r.quarter_rectified === q && r.tat_follow === "InTAT").length), backgroundColor: PALETTE[5] },
      { label: "Out of TAT", data: quarters.map(q => closed.filter(r => r.quarter_rectified === q && r.tat_follow === "OutTAT").length), backgroundColor: PALETTE[3] } ] },
    options: { scales: { x: { stacked: true }, y: { stacked: true } } } });

  makeChart("ch-tat-trend", { type: "line",
    data: { labels: quarters, datasets: [
      { label: "Overall", data: quarters.map(q => avg(closed.filter(r => r.quarter_rectified === q), r => r.rectification_time)), borderColor: PALETTE[0], tension: .3 },
      { label: "Channelplay", data: quarters.map(q => avg(closed.filter(r => r.quarter_rectified === q && r.responsibility === "Channelplay"), r => r.rectification_time)), borderColor: PALETTE[1], tension: .3 },
      { label: "ASUS/RV", data: quarters.map(q => avg(closed.filter(r => r.quarter_rectified === q && r.responsibility !== "Channelplay"), r => r.rectification_time)), borderColor: PALETTE[2], tension: .3 } ] } });

  const oaRows = bucketNames.map((bn, i) => {
    const inB = open.filter(t => openBucket(openAgeDays(t)) === bn);
    return { bn, total: openCounts[i],
      cp: inB.filter(t => t.responsibility === "Channelplay").length,
      rv: inB.filter(t => t.responsibility !== "Channelplay").length,
      share: pct(openCounts[i], open.length) };
  }).filter(r => r.total > 0);
  S.agg.openAgeing = { headers: ["Ageing Bucket","Open Tickets","Channelplay","ASUS/RV","Share %"],
    rows: oaRows.map(r => [r.bn, r.total, r.cp, r.rv, r.share]) };
  renderTable("tbl-open-ageing", S.agg.openAgeing.headers, oaRows.map(r => ({
    cells: [r.bn, fmt(r.total), fmt(r.cp), fmt(r.rv), r.share], numCols: [1,2,3,4],
    onClick: () => drill({ final: "Open" }, r.bn) })),
    { totals: ["Total", fmt(open.length), fmt(open.filter(t => t.responsibility === "Channelplay").length), fmt(open.filter(t => t.responsibility !== "Channelplay").length), "100%"] });
}

// ============================================================
// DATA EXPLORER
// ============================================================
const EXPLORER_COLS = [
  ["ticket_id","Ticket ID"],["issue_raised_date","Raised"],["region","Region"],["branch","Branch"],
  ["store_name","Store"],["city","City"],["city_classification","Tier"],["issue_category","Issue Category"],
  ["budget_category","Budget"],["status","Stage"],["final_status","Open/Closed"],["responsibility","Resp."],
  ["rectification_date","Rectified"],["rectification_time","TAT (days)"],["tat_follow","TAT Flag"],
];
function renderExplorer() {
  const rows = getFiltered();
  $("explorer-count").textContent = fmt(rows.length) + " tickets";
  const pages = Math.max(1, Math.ceil(rows.length / S.pageSize));
  if (S.page >= pages) S.page = pages - 1;
  const slice = rows.slice(S.page * S.pageSize, (S.page + 1) * S.pageSize);
  $("pg-info").textContent = `Page ${S.page + 1} of ${pages}`;
  renderTable("tbl-explorer", EXPLORER_COLS.map(c => c[1]), slice.map(t => ({
    cells: EXPLORER_COLS.map(([k]) => {
      if (k === "final_status") return `<span class="pill ${t[k] === "Open" ? "open" : "closed"}">${esc(t[k] || "-")}</span>`;
      if (k.endsWith("_date")) return fmtDate(t[k]);
      return esc(t[k] == null ? "-" : t[k]);
    }),
    numCols: [13], onClick: () => openDrawer(t) })));
}

function openDrawer(t) {
  $("drawer-title").textContent = "Ticket #" + t.ticket_id;
  const skip = new Set(["extra", "updated_at"]);
  let html = "";
  for (const [k, v] of Object.entries(t)) {
    if (skip.has(k) || v == null || v === "") continue;
    html += `<div class="field"><label>${esc(k.replace(/_/g, " "))}</label>${esc(k.endsWith("_date") ? fmtDate(v) : v)}</div>`;
  }
  if (t.extra && Object.keys(t.extra).length) {
    html += `<div class="field"><label>— additional columns —</label></div>`;
    for (const [k, v] of Object.entries(t.extra)) html += `<div class="field"><label>${esc(k)}</label>${esc(v)}</div>`;
  }
  $("drawer-body").innerHTML = html;
  $("drawer").classList.remove("hidden"); $("drawer-overlay").classList.remove("hidden");
}
function closeDrawer() { $("drawer").classList.add("hidden"); $("drawer-overlay").classList.add("hidden"); }

// ============================================================
// EXPORT
// ============================================================
function exportRaw(kind) {
  const rows = getFiltered();
  if (!rows.length) return toast("Nothing to export with current filters");
  const flat = rows.map(t => {
    const { extra, updated_at, ...main } = t;
    return { ...main, ...(extra || {}) };
  });
  const stamp = new Date().toISOString().slice(0, 10);
  const name = `asus_tickets_filtered_${stamp}`;
  if (kind === "xlsx") {
    const ws = XLSX.utils.json_to_sheet(flat);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tickets");
    // filter summary sheet
    const fs = Object.entries(F).filter(([, v]) => v).map(([k, v]) => ({ filter: k, value: v }));
    if (fs.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fs), "Applied Filters");
    XLSX.writeFile(wb, name + ".xlsx");
  } else {
    const ws = XLSX.utils.json_to_sheet(flat);
    downloadText(name + ".csv", XLSX.utils.sheet_to_csv(ws));
  }
  toast(`Exported ${fmt(rows.length)} tickets (${kind.toUpperCase()})`);
}

function exportAgg(key) {
  const a = S.agg[key];
  if (!a) return toast("Open this view first, then export");
  const lines = [a.headers.join(",")].concat(a.rows.map(r => r.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")));
  downloadText(`asus_${key}_${new Date().toISOString().slice(0, 10)}.csv`, lines.join("\n"));
}

function downloadText(name, text) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([text], { type: "text/csv" }));
  a.download = name; a.click(); URL.revokeObjectURL(a.href);
}

// ============================================================
// generic sortable table renderer
// ============================================================
function renderTable(id, headers, rowDefs, opts = {}) {
  const tbl = $(id);
  const sortState = S.sort[id] || { col: -1, dir: 1 };
  if (sortState.col >= 0) {
    rowDefs = [...rowDefs].sort((a, b) => {
      const av = stripHtml(a.cells[sortState.col]), bv = stripHtml(b.cells[sortState.col]);
      const an = parseFloat(String(av).replace(/[,%]/g, "")), bn = parseFloat(String(bv).replace(/[,%]/g, ""));
      const cmp = (!isNaN(an) && !isNaN(bn)) ? an - bn : String(av).localeCompare(String(bv));
      return cmp * sortState.dir;
    });
  }
  let html = "<thead><tr>" + headers.map((h, i) =>
    `<th data-col="${i}">${esc(h)}${sortState.col === i ? (sortState.dir > 0 ? " ▲" : " ▼") : ""}</th>`).join("") + "</tr></thead><tbody>";
  for (const r of rowDefs) {
    html += "<tr>" + r.cells.map((c, i) => `<td class="${(r.numCols || []).includes(i) ? "num" : ""}">${c}</td>`).join("") + "</tr>";
  }
  if (opts.totals) html += "<tr class='total-row'>" + opts.totals.map((c, i) => `<td class="${i ? "num" : ""}">${c}</td>`).join("") + "</tr>";
  html += "</tbody>";
  tbl.innerHTML = html;
  tbl.querySelectorAll("th").forEach(th => th.onclick = () => {
    const col = Number(th.dataset.col);
    const st = S.sort[id] || { col: -1, dir: 1 };
    S.sort[id] = { col, dir: st.col === col ? -st.dir : -1 };
    renderActive(true);
  });
  tbl.querySelectorAll("tbody tr").forEach((tr, i) => {
    if (rowDefs[i] && rowDefs[i].onClick) tr.onclick = () => rowDefs[i].onClick();
  });
}
function stripHtml(s) { const d = document.createElement("div"); d.innerHTML = s; return d.textContent; }

// ---------------- misc ----------------
function toast(msg, ms = 2600) {
  const t = $("toast"); t.textContent = msg; t.classList.remove("hidden");
  clearTimeout(t._h); t._h = setTimeout(() => t.classList.add("hidden"), ms);
}
function loading(on, msg) {
  $("loading").classList.toggle("hidden", !on);
  if (msg) $("loading-msg").textContent = msg;
}
function initTheme() {
  const saved = localStorage.getItem("asus_theme");
  if (saved) document.documentElement.setAttribute("data-theme", saved);
}
function toggleTheme() {
  const cur = document.documentElement.getAttribute("data-theme") === "dark" ? "" : "dark";
  cur ? document.documentElement.setAttribute("data-theme", cur) : document.documentElement.removeAttribute("data-theme");
  localStorage.setItem("asus_theme", cur);
  S.dirty = true; renderActive(true); // re-render charts with new colors
}

// expose for upload.js
window.App = { S, loadData, toast, loading, fmt, esc, fmtDate };
