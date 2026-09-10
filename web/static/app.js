"use strict";
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const CRIT = window.CRITICAL_DAYS;

let STATE = {
  view: "default",
  products: [],        // selected Prodgroup3
  filters: {},         // {col: [values]}
  scope: "all",
  reconProducts: [],   // last reconciliation products
  atmfProducts: [],    // ATMf product-name option list
  picks: {},           // lot -> {product, pi} ticked for ATMf
};

/* ---------- helpers ---------- */
function toast(msg, ms = 2600) {
  const t = $("#toast"); t.textContent = msg; t.classList.add("show");
  clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove("show"), ms);
}
function overlay(on, msg = "") { const o = $("#overlay"); $("#overlayMsg").textContent = msg; o.classList.toggle("hidden", !on); }
async function api(url, opts) {
  const res = await fetch(url, opts);
  const data = await res.json().catch(() => ({ ok: false, error: "bad json" }));
  if (!res.ok || data.ok === false) throw new Error(data.error || ("HTTP " + res.status));
  return data;
}
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

/* ---------- filters ---------- */
async function loadOptions() {
  const data = await api("/api/options?view=" + STATE.view);
  const { options, preset, all_products } = data;
  // Products as chips.
  const prods = options["Prodgroup3"] || [];
  const presetProds = new Set(all_products ? prods : (preset["Prodgroup3"] || []));
  STATE.products = prods.filter(p => presetProds.has(p));
  const chips = $("#productChips"); chips.innerHTML = "";
  prods.forEach(p => {
    const c = document.createElement("span");
    c.className = "chip" + (STATE.products.includes(p) ? " on" : "");
    c.textContent = p;
    c.onclick = () => { c.classList.toggle("on"); syncProducts(); };
    chips.appendChild(c);
  });
  // Other filter columns as checkbox lists.
  const grid = $("#filterGrid"); grid.innerHTML = "";
  STATE.filters = {};
  ["Operation", "Mgr_Name", "Department"].forEach(col => {
    const vals = options[col] || [];
    const chosen = new Set(preset[col] || []);
    STATE.filters[col] = vals.filter(v => chosen.has(v));
    const box = document.createElement("div");
    box.innerHTML = `<label>${col}</label>`;
    const multi = document.createElement("div"); multi.className = "multi";
    vals.forEach(v => {
      const id = `f_${col}_${v}`.replace(/\W/g, "_");
      const lab = document.createElement("label");
      lab.innerHTML = `<input type="checkbox" ${chosen.has(v) ? "checked" : ""}> ${esc(v)}`;
      lab.querySelector("input").onchange = e => {
        const set = new Set(STATE.filters[col]);
        e.target.checked ? set.add(v) : set.delete(v);
        STATE.filters[col] = [...set];
      };
      multi.appendChild(lab);
    });
    box.appendChild(multi); grid.appendChild(box);
  });
}
function syncProducts() {
  STATE.products = $$("#productChips .chip.on").map(c => c.textContent);
}

/* ---------- summary ---------- */
async function loadSummary() {
  overlay(true, "Loading summary…");
  try {
    const data = await api("/api/summary", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ products: STATE.products, filters: STATE.filters }),
    });
    renderMetrics(data);
    renderProdSummary(data.products);
  } catch (e) { toast("Summary failed: " + e.message); }
  finally { overlay(false); }
}
function renderMetrics(d) {
  const m = $("#metrics");
  const totalPPV = d.products.reduce((a, p) => a + p.ppv_lots, 0);
  const totalCls = d.products.reduce((a, p) => a + p.class_lots, 0);
  const totalCrit = d.products.reduce((a, p) => a + p.critical, 0);
  m.innerHTML = [
    ["Products", d.n_products], ["Total rows", d.total_rows],
    ["PPV lots", totalPPV], ["Class lots", totalCls],
  ].map(([l, v]) => `<div class="metric"><div class="v">${v}</div><div class="l">${l}</div></div>`).join("")
    + `<div class="metric bad"><div class="v">${totalCrit}</div><div class="l">Critical &gt;${CRIT}d</div></div>`;
}
function renderProdSummary(prods) {
  const tb = $("#prodSummary tbody"); tb.innerHTML = "";
  prods.forEach(p => {
    const tr = document.createElement("tr"); tr.className = "clickable";
    tr.innerHTML = `<td>${esc(p.product)}</td><td class="num">${p.lots}</td>
      <td class="num">${p.unit}</td><td class="num">${p.ppv_lots}</td>
      <td class="num">${p.class_lots}</td><td class="num">${p.critical}</td>
      <td class="linkcell"><a href="#prod-${p.slug}">Open ↓</a></td>`;
    tr.onclick = ev => { if (ev.target.tagName !== "A") jumpTo(p.slug); };
    tb.appendChild(tr);
  });
}
function jumpTo(slug) {
  const prods = STATE.reconProducts || [];
  const idx = prods.findIndex(p => p.slug === slug);
  if (idx >= 0) {
    showProductDetail(idx);
    const rc = document.getElementById("reconCard");
    if (rc) rc.scrollIntoView({ behavior: "smooth", block: "start" });
  } else {
    toast("Run/refresh reconciliation first to see product detail.");
  }
}

/* ---------- reconciliation ---------- */
async function runReconcile() {
  overlay(true, "Querying RUPS…");
  try {
    const data = await api("/api/reconcile", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ products: STATE.products, filters: STATE.filters, scope: STATE.scope }),
    });
    STATE.reconProducts = data.products || [];
    STATE.picks = {};
    renderReconMetrics(data);
    renderReconProducts(data.products || []);
    updateAtmfPreview();
    if (data.message) {
      $("#reconStatus").textContent = data.message;
      toast(data.message);
    } else {
      $("#reconStatus").textContent = `Searched ${data.lots} lot(s), ${data.records} RUPS row(s).`;
    }
  } catch (e) { toast("Reconcile failed: " + e.message); }
  finally { overlay(false); }
}
function renderReconMetrics(d) {
  const m = d.metrics || {};
  $("#reconMetrics").innerHTML = [
    ["PPV units", m.ppv_units || 0, ""], ["Class units", m.class_units || 0, ""],
    ["Lots not found", m.not_found || 0, m.not_found ? "bad" : ""],
    ["Qty mismatches", m.mismatch || 0, m.mismatch ? "bad" : ""],
  ].map(([l, v, c]) => `<div class="metric ${c}"><div class="v">${v}</div><div class="l">${l}</div></div>`).join("");
}
function renderReconProducts(prods) {
  const wrap = $("#reconProducts"); wrap.innerHTML = "";
  const bar = $("#reconToolbar");
  if (!prods.length) {
    wrap.innerHTML = `<p class="muted">No lots to reconcile.</p>`;
    if (bar) bar.classList.add("hidden");
    return;
  }
  if (bar) bar.classList.remove("hidden");
  // Sort like the top summary table: most issues first (mismatch desc, then lots desc).
  prods = prods.slice().sort((a, b) =>
    (b.mismatch - a.mismatch) || (b.n_lots - a.n_lots) ||
    String(a.product).localeCompare(String(b.product)));
  STATE.reconProducts = prods;

  // Two-pane layout: left = compact product list, right = one product's detail.
  const pane = document.createElement("div"); pane.className = "recon-split";
  const nav = document.createElement("div"); nav.className = "prod-nav";
  const detail = document.createElement("div"); detail.className = "prod-detail";
  detail.id = "prodDetail";
  pane.appendChild(nav); pane.appendChild(detail);
  wrap.appendChild(pane);

  prods.forEach((p, idx) => {
    const item = document.createElement("button");
    item.className = "prod-nav-item" + (p.mismatch ? " has-mis" : "");
    item.dataset.idx = idx;
    item.innerHTML = `<span class="pn-name">${esc(p.product)}</span>
      <span class="pn-meta"><span class="pn-count">${p.n_lots}</span>
      ${p.mismatch ? `<span class="badge mis">⚠${p.mismatch}</span>` : ""}</span>`;
    item.onclick = () => showProductDetail(idx);
    nav.appendChild(item);
  });

  // Show the first product with a mismatch, else the first product.
  const first = prods.findIndex(p => p.mismatch > 0);
  showProductDetail(first >= 0 ? first : 0);
}

function showProductDetail(idx) {
  const prods = STATE.reconProducts || [];
  const p = prods[idx];
  if (!p) return;
  $$("#reconProducts .prod-nav-item").forEach(el =>
    el.classList.toggle("active", +el.dataset.idx === idx));
  const detail = $("#prodDetail"); if (!detail) return;
  detail.innerHTML = "";
  const head = document.createElement("div"); head.className = "prod-head";
  head.innerHTML = `<span class="prod-name">📦 ${esc(p.product)}</span>
    <span class="muted">${p.n_lots} lot(s)</span>
    <span class="badge ppv">PPV ${p.n_ppv}</span>
    <span class="badge cls">Class ${p.n_class}</span>
    ${p.n_eng ? `<span class="badge eng">Eng ${p.n_eng}</span>` : ""}
    ${p.mismatch ? `<span class="badge mis">⚠ ${p.mismatch} mismatch</span>` : ""}`;
  const selBtn = document.createElement("button");
  selBtn.className = "btn btn-sm"; selBtn.textContent = "Select all in this product";
  selBtn.onclick = () => selectLots("all", detail);
  head.appendChild(selBtn);
  detail.appendChild(head);

  const tabs = document.createElement("div"); tabs.className = "pi-tabs";
  const panes = document.createElement("div");
  p.groups.forEach((g, i) => {
    const tab = document.createElement("div");
    tab.className = "pi-tab" + (i === 0 ? " active" : "");
    tab.textContent = `${g.pi} (${g.lots.length})`;
    const pn = document.createElement("div");
    pn.style.display = i === 0 ? "block" : "none";
    pn.appendChild(lotTable(g.lots, p.product));
    tab.onclick = () => {
      $$(".pi-tab", tabs).forEach(t => t.classList.remove("active"));
      $$("div", panes).forEach(x => x.style.display = "none");
      tab.classList.add("active"); pn.style.display = "block";
    };
    tabs.appendChild(tab); panes.appendChild(pn);
  });
  detail.appendChild(tabs); detail.appendChild(panes);
}
function lotTable(lots, product) {
  const tbl = document.createElement("table"); tbl.className = "tbl lot-tbl";
  const flag = s => s === "MISMATCH" ? "⚠" : s === "Not found" ? "❓" : s === "N/A" ? "–" : "✓";
  tbl.innerHTML = `<thead><tr>
    <th><input type="checkbox" class="hdr-check" title="Toggle all in this tab"></th>
    <th>!</th><th>Lot</th>
    <th class="num">EIMS</th><th class="num">RUPS</th><th class="num">Last used (d)</th>
    <th>Operation</th><th>PI Dispose</th></tr></thead>`;
  const tb = document.createElement("tbody");
  lots.forEach(r => {
    const tr = document.createElement("tr");
    const mis = r.status === "MISMATCH";
    const checked = STATE.picks[r.lot] ? "checked" : "";
    tr.innerHTML = `<td><input type="checkbox" ${checked} data-lot="${esc(r.lot)}"></td>
      <td class="${mis ? "flag-mis" : ""}">${flag(r.status)}</td>
      <td>${esc(r.lot)}</td>
      <td class="num ${mis ? "mismatch" : ""}">${r.eims_qty == null ? "" : r.eims_qty}</td>
      <td class="num ${mis ? "mismatch" : ""}">${r.rups_qty}</td>
      <td class="num">${r.last_used == null ? "" : r.last_used}</td>
      <td class="${r.op_conflict ? "conflict" : ""}">${esc(r.operation)}</td>
      <td>${esc(r.pi_dispose)}</td>`;
    const cb = tr.querySelector("input");
    cb._rec = { lot: r.lot, product, pi: r.pi_dispose };
    cb.onchange = ev => applyCheck(cb, ev);
    // clicking anywhere on the row toggles the checkbox
    tr.onclick = ev => {
      if (ev.target.tagName === "INPUT") return;
      cb.checked = !cb.checked; applyCheck(cb, ev);
    };
    tb.appendChild(tr);
  });
  tbl.appendChild(tb);
  // Header checkbox toggles the whole tab.
  const hdr = tbl.querySelector(".hdr-check");
  hdr.onchange = () => {
    $$("tbody input[type=checkbox]", tbl).forEach(c => {
      if (c.checked !== hdr.checked) { c.checked = hdr.checked; applyCheck(c); }
    });
  };
  return tbl;
}

// Track last-clicked checkbox per table for shift-range selection.
let _lastCb = null;
function applyCheck(cb, ev) {
  const rec = cb._rec;
  if (!rec) return;
  // Shift-click range within the same tbody.
  if (ev && ev.shiftKey && _lastCb && _lastCb !== cb) {
    const body = cb.closest("tbody");
    if (body && _lastCb.closest("tbody") === body) {
      const all = $$("input[type=checkbox]", body);
      const a = all.indexOf(_lastCb), b = all.indexOf(cb);
      const [lo, hi] = [Math.min(a, b), Math.max(a, b)];
      for (let i = lo; i <= hi; i++) {
        all[i].checked = cb.checked;
        const rr = all[i]._rec;
        if (rr) { cb.checked ? (STATE.picks[rr.lot] = rr) : delete STATE.picks[rr.lot]; }
      }
    }
  }
  if (cb.checked) STATE.picks[rec.lot] = rec; else delete STATE.picks[rec.lot];
  _lastCb = cb;
  refreshSelection();
}

function selectLots(mode, root) {
  if (!root) {
    // Global toolbar: operate on ALL products (only one is rendered in the DOM),
    // so drive selection from the data, then sync any visible checkboxes.
    (STATE.reconProducts || []).forEach(p =>
      p.groups.forEach(g => g.lots.forEach(r => {
        let want;
        if (mode === "all") want = true;
        else if (mode === "clear") want = false;
        else if (mode === "mismatch") want = (r.status === "MISMATCH") || (STATE.picks[r.lot] != null);
        if (mode === "mismatch") want = (r.status === "MISMATCH");
        if (want) STATE.picks[r.lot] = { lot: r.lot, product: p.product, pi: r.pi_dispose };
        else delete STATE.picks[r.lot];
      })));
    $$("#reconProducts tbody input[type=checkbox]").forEach(cb => {
      const rec = cb._rec; if (rec) cb.checked = STATE.picks[rec.lot] != null;
    });
    refreshSelection();
    return;
  }
  $$("tbody input[type=checkbox]", root).forEach(cb => {
    const rec = cb._rec; if (!rec) return;
    let want = cb.checked;
    if (mode === "all") want = true;
    else if (mode === "clear") want = false;
    else if (mode === "mismatch") want = !!cb.closest("tr").querySelector("td.mismatch");
    cb.checked = want;
    if (want) STATE.picks[rec.lot] = rec; else delete STATE.picks[rec.lot];
  });
  refreshSelection();
}

function refreshSelection() {
  const n = Object.keys(STATE.picks).length;
  const c = $("#selCount"); if (c) c.textContent = n + " selected";
  // Sync header checkboxes.
  $$("#reconProducts table.lot-tbl").forEach(tbl => {
    const hdr = tbl.querySelector(".hdr-check");
    const boxes = $$("tbody input[type=checkbox]", tbl);
    if (hdr) hdr.checked = boxes.length > 0 && boxes.every(b => b.checked);
  });
  updateAtmfPreview();
}

/* ---------- ATMf ---------- */
async function updateAtmfPreview() {
  await guessProducts();
  const picks = STATE.picks;
  const testMode = $("#atmfTest").value;
  // Group by (product, testArea).
  const groups = {};
  Object.entries(picks).forEach(([lot, info]) => {
    const area = testMode === "auto" ? (info.pi === "PPV" ? "PPV" : "Class") : testMode;
    const key = info.product + "||" + area;
    if (!groups[key]) groups[key] = { product: info.product, area, lots: [] };
    groups[key].lots.push(lot);
  });
  const list = Object.values(groups);
  // Guess ATMf product names.
  const tb = $("#atmfPreview tbody"); tb.innerHTML = "";
  const prodMap = STATE._guessMap || {};
  let allMapped = list.length > 0;
  list.forEach((g, i) => {
    const name = prodMap[g.product] || "";
    if (!name) allMapped = false;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${reqSelect(i)}</td><td>${esc(g.product)}</td>
      <td>${nameSelect(i, name)}</td><td>${esc(g.area)}</td>
      <td>${esc(g.lots.join(", "))}</td><td class="num">${g.lots.length}</td>`;
    tb.appendChild(tr);
    g._name = name;
  });
  STATE._atmfGroups = list;
  $("#btnSubmitAtmf").disabled = !allMapped;
}
function reqSelect(i) {
  return `<select data-req="${i}">${INTMS.requests.map(r => `<option>${esc(r)}</option>`).join("")}</select>`;
}
function nameSelect(i, sel) {
  const opts = STATE.atmfProducts.map(p => `<option ${p === sel ? "selected" : ""}>${esc(p)}</option>`).join("");
  return `<select data-name="${i}"><option value="">(pick)</option>${opts}</select>`;
}
async function guessProducts() {
  const prodgroups = [...new Set(Object.values(STATE.picks).map(p => p.product))];
  if (!prodgroups.length) { STATE._guessMap = {}; return; }
  try {
    const data = await api("/api/atmf/guess", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prodgroups }),
    });
    STATE._guessMap = data.map || {};
  } catch (e) { STATE._guessMap = {}; }
}
async function submitAtmf() {
  const list = STATE._atmfGroups || [];
  const tickets = list.map((g, i) => {
    const reqSel = $(`#atmfPreview select[data-req="${i}"]`);
    const nameSel = $(`#atmfPreview select[data-name="${i}"]`);
    return {
      product_name: nameSel ? nameSel.value : g._name,
      test_area: g.area,
      request: reqSel ? reqSel.value : INTMS.requests[0],
      lots: g.lots,
      customer: $("#atmfCustomer").value, payer: $("#atmfPayer").value,
      stage: $("#atmfStage").value, factory: $("#atmfFactory").value,
      shipping: $("#atmfShipping").value, nfo: $("#atmfNfo").value,
    };
  });
  if (tickets.some(t => !t.product_name)) { toast("Pick a Product Name for every row."); return; }
  overlay(true, "Submitting tickets…");
  try {
    const data = await api("/api/atmf/submit", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ tickets }),
    });
    const box = $("#atmfResults"); box.innerHTML = "";
    data.results.forEach(r => {
      const d = document.createElement("div");
      d.className = "res " + (r.ok ? "ok" : "err");
      if (r.ok) d.innerHTML = `✓ ${esc(r.lots.join(", "))} (${esc(r.test_area)}) → ticket #${r.ticket_id}
        ${r.filled ? "" : "(fill-in not prefilled)"} <a href="${r.url}" target="_blank">${r.url}</a>`;
      else d.innerHTML = `✗ ${esc(r.lots.join(", "))} failed: ${esc(r.detail)}`;
      box.appendChild(d);
    });
  } catch (e) { toast("Submit failed: " + e.message); }
  finally { overlay(false); }
}

/* ---------- email ---------- */
async function previewEmail() {
  overlay(true, "Building report…");
  try {
    const data = await api("/api/email/preview", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
    if (!$("#emailSubject").value) $("#emailSubject").value = data.subject;
    $("#emailPreview").innerHTML = data.html;
  } catch (e) { toast("Preview failed: " + e.message); }
  finally { overlay(false); }
}
async function openOutlook() {
  overlay(true, "Opening Outlook…");
  try {
    await api("/api/email/open", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ to: $("#emailTo").value, cc: $("#emailCc").value, subject: $("#emailSubject").value }),
    });
    toast("Outlook draft opened — review and send.");
  } catch (e) { toast("Outlook failed: " + e.message); }
  finally { overlay(false); }
}

/* ---------- theme ---------- */
function applyTheme(t) {
  document.documentElement.setAttribute("data-theme", t);
  const tt = $("#themeToggle .tt-icon");
  if (tt) tt.textContent = t === "dark" ? "🌙" : "☀️";
  try { localStorage.setItem("rm_theme", t); } catch (e) {}
}
function initTheme() {
  let t = "dark";
  try { t = localStorage.getItem("rm_theme") || "dark"; } catch (e) {}
  applyTheme(t);
  const btn = $("#themeToggle");
  if (btn) btn.onclick = () =>
    applyTheme(document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark");
}

/* ---------- wiring ---------- */
async function init() {
  initTheme();
  const p = await api("/api/atmf/products").catch(() => ({ products: [] }));
  STATE.atmfProducts = p.products || [];

  $("#btnRefresh").onclick = async () => {
    overlay(true, "Downloading EIMS…");
    try {
      const d = await api("/api/refresh", { method: "POST" });
      $("#lastUpdate").textContent = "Last update: " + d.last_update;
      window.HAS_EIMS = true;
      await loadOptions(); await loadSummary(); await runReconcile();
      toast("EIMS refreshed.");
    } catch (e) { toast("Refresh failed: " + e.message); }
    finally { overlay(false); }
  };

  $$("#viewSeg button").forEach(b => b.onclick = async () => {
    $$("#viewSeg button").forEach(x => x.classList.remove("active"));
    b.classList.add("active"); STATE.view = b.dataset.view;
    await loadOptions(); await loadSummary(); await runReconcile();
  });
  $$("#scopeSeg button").forEach(b => b.onclick = async () => {
    $$("#scopeSeg button").forEach(x => x.classList.remove("active"));
    b.classList.add("active"); STATE.scope = b.dataset.scope;
    await runReconcile();
  });

  $("#btnApply").onclick = async () => { syncProducts(); await loadSummary(); await runReconcile(); };
  $("#btnSubmitAtmf").onclick = submitAtmf;
  $("#atmfTest").onchange = updateAtmfPreview;
  $("#btnPreview").onclick = previewEmail;
  $("#btnOpenOutlook").onclick = openOutlook;
  $$("#reconToolbar button[data-sel]").forEach(b =>
    b.onclick = () => selectLots(b.dataset.sel));

  if (window.HAS_EIMS) {
    await loadOptions(); await loadSummary(); await runReconcile();
  } else {
    toast("No EIMS file yet — click Refresh EIMS.");
  }
}
document.addEventListener("DOMContentLoaded", init);
