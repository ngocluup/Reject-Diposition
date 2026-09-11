"use strict";
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const CRIT = window.CRITICAL_DAYS;

let STATE = {
  view: "default",
  products: [],        // selected Prodgroup3
  allProducts: [],     // every Prodgroup3 available in the current view
  defaultProducts: [], // the preset selection, used by the Reset button
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
  const prods = options["Prodgroup3"] || [];
  const presetProds = new Set(all_products ? prods : (preset["Prodgroup3"] || []));
  STATE.allProducts = prods;
  STATE.defaultProducts = prods.filter(p => presetProds.has(p));
  STATE.products = STATE.defaultProducts.slice();
  renderProductPicker();

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

/* Group codes into product families so 150+ chips stay readable.
   ADLN / ADLP282 / ADL282IOTG -> "ADL";  PP-* -> "PP (pre-production)". */
function productFamily(code) {
  const s = String(code).toUpperCase();
  if (s.startsWith("PP-")) return "PP · pre-production";
  const m = s.match(/^[A-Z]+/);
  if (!m) return "0-9";
  const letters = m[0];
  // 3 letters is the usual Intel family prefix (ADL, ARL, RPL, GNR…).
  return letters.length >= 3 ? letters.slice(0, 3) : letters;
}

function renderProductPicker() {
  const wrap = $("#productChips"); wrap.innerHTML = "";
  const selected = new Set(STATE.products);

  // Bucket by family, keeping one-off codes together under "Other".
  const fams = {};
  STATE.allProducts.forEach(p => {
    const f = productFamily(p);
    (fams[f] = fams[f] || []).push(p);
  });
  const singles = [];
  Object.keys(fams).forEach(f => {
    if (fams[f].length === 1 && f !== "PP · pre-production") {
      singles.push(fams[f][0]); delete fams[f];
    }
  });
  if (singles.length) fams["Other"] = singles.sort();

  // Families with a selected product first, then biggest, then alphabetical.
  const order = Object.keys(fams).sort((a, b) => {
    const sa = fams[a].some(p => selected.has(p)) ? 0 : 1;
    const sb = fams[b].some(p => selected.has(p)) ? 0 : 1;
    if (sa !== sb) return sa - sb;
    if (a === "Other") return 1;
    if (b === "Other") return -1;
    return fams[b].length - fams[a].length || a.localeCompare(b);
  });

  order.forEach(fam => {
    const items = fams[fam].slice().sort();
    const nOn = items.filter(p => selected.has(p)).length;

    const grp = document.createElement("div");
    grp.className = "chip-group" + (nOn ? " has-on" : "");
    grp.dataset.family = fam;

    const head = document.createElement("div"); head.className = "cg-head";
    head.innerHTML = `<span class="cg-name">${esc(fam)}</span>
      <span class="cg-count"><b class="cg-on">${nOn}</b>/${items.length}</span>`;
    // Clicking the family header toggles the whole family.
    head.title = "Click to select / clear this family";
    head.onclick = () => {
      const set = new Set(STATE.products);
      const allOn = items.every(p => set.has(p));
      items.forEach(p => allOn ? set.delete(p) : set.add(p));
      STATE.products = STATE.allProducts.filter(p => set.has(p));
      renderProductPicker();
      applyProductSearch($("#prodSearch").value);
    };
    grp.appendChild(head);

    const row = document.createElement("div"); row.className = "cg-chips";
    items.forEach(p => {
      const c = document.createElement("span");
      c.className = "chip" + (selected.has(p) ? " on" : "");
      c.textContent = p;
      c.dataset.code = p.toUpperCase();
      c.onclick = () => {
        c.classList.toggle("on");
        syncProducts();
        updateProductCounts();
      };
      row.appendChild(c);
    });
    grp.appendChild(row);
    wrap.appendChild(grp);
  });
  updateProductCounts();
}

function syncProducts() {
  STATE.products = $$("#productChips .chip.on").map(c => c.textContent);
}

function updateProductCounts() {
  const sel = new Set(STATE.products);
  const c = $("#prodCount");
  if (c) c.textContent = `${sel.size} / ${STATE.allProducts.length}`;
  $$("#productChips .chip-group").forEach(g => {
    const chips = $$(".chip", g);
    const on = chips.filter(x => x.classList.contains("on")).length;
    const b = g.querySelector(".cg-on");
    if (b) b.textContent = on;
    g.classList.toggle("has-on", on > 0);
  });
}

function applyProductSearch(q) {
  const term = String(q || "").trim().toUpperCase();
  $$("#productChips .chip-group").forEach(g => {
    // Typing a family name (e.g. "ADL") keeps the whole family visible.
    const famHit = !!term && g.dataset.family.toUpperCase().includes(term);
    let visible = 0;
    $$(".chip", g).forEach(c => {
      const hit = !term || famHit || c.dataset.code.includes(term);
      c.classList.toggle("hide", !hit);
      if (hit) visible++;
    });
    g.classList.toggle("hide", visible === 0);
  });
}

function setProducts(list) {
  STATE.products = list.slice();
  renderProductPicker();
  applyProductSearch($("#prodSearch").value);
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
      <td class="num">${p.class_lots}</td><td class="num">${p.critical}</td>`;
    tr.onclick = () => jumpTo(p.slug);
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
    updateActionBar(0);
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
    item.className = "prod-nav-item";
    item.dataset.idx = idx;
    item.innerHTML = `<span class="pn-name">${esc(p.product)}</span>
      <span class="pn-meta"><span class="pn-count">${p.n_lots}</span></span>`;
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
    ${p.n_eng ? `<span class="badge eng">Eng ${p.n_eng}</span>` : ""}`;
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
    // Scroll the lot list inside a fixed-height box so the page stays short.
    const box = document.createElement("div"); box.className = "lot-scroll";
    box.appendChild(lotTable(g.lots, p.product));
    pn.appendChild(box);
    pn.appendChild(exportBar(g.lots, p.product, g.pi));
    tab.onclick = () => {
      $$(".pi-tab", tabs).forEach(t => t.classList.remove("active"));
      $$("div", panes).forEach(x => { if (x.parentElement === panes) x.style.display = "none"; });
      tab.classList.add("active"); pn.style.display = "block";
    };
    tabs.appendChild(tab); panes.appendChild(pn);
  });
  detail.appendChild(tabs); detail.appendChild(panes);
}

/* ---------- export a lot table ---------- */
const EXPORT_COLS = [
  ["lot", "Lot"], ["product", "Product"], ["product_id", "Product ID"],
  ["eims_qty", "EIMS_Qty"], ["rups_qty", "RUPS_Qty"], ["last_used", "Last_Used_Days"],
  ["operation", "Operation"], ["pi_dispose", "PI_Dispose"], ["status", "Status"],
];

function exportBar(lots, product, pi) {
  const bar = document.createElement("div"); bar.className = "export-bar";
  const info = document.createElement("span"); info.className = "muted";
  bar.appendChild(info);

  const mk = (label, title, fn) => {
    const b = document.createElement("button");
    b.className = "btn btn-sm btn-ghost"; b.textContent = label; b.title = title;
    b.onclick = () => {
      // Export the ticked rows if there are any, otherwise the whole tab.
      const picked = lots.filter(r => STATE.picks[r.lot]);
      const rows = picked.length ? picked : lots;
      fn(rows, `${slugName(product)}_${slugName(pi)}_${stamp()}`);
      toast(`Exported ${rows.length} row(s)${picked.length ? " (selected only)" : ""}.`);
    };
    return b;
  };
  bar.appendChild(mk("⬇ CSV", "Download this table as CSV (opens in Excel)", downloadCsv));
  bar.appendChild(mk("⬇ Excel", "Download this table as an Excel file", downloadXls));
  bar.appendChild(mk("⧉ Copy", "Copy the table to the clipboard", copyRows));

  // Keep the hint in sync with the current selection.
  const sync = () => {
    const n = lots.filter(r => STATE.picks[r.lot]).length;
    info.textContent = n ? `Exports the ${n} selected row(s)` : `Exports all ${lots.length} row(s)`;
  };
  sync();
  bar._sync = sync;
  return bar;
}

function rowsToMatrix(rows) {
  const head = EXPORT_COLS.map(([, h]) => h);
  const body = rows.map(r => EXPORT_COLS.map(([k]) => {
    const v = r[k];
    return v == null ? "" : String(v);
  }));
  return [head, ...body];
}
function slugName(s) {
  return String(s || "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "") || "data";
}
function stamp() {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
}
function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function downloadCsv(rows, base) {
  const csv = rowsToMatrix(rows).map(r =>
    r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\r\n");
  // UTF-8 BOM so Excel opens it with the right encoding.
  saveBlob(new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" }), base + ".csv");
}
function downloadXls(rows, base) {
  // SpreadsheetML table - Excel opens this natively, no library needed.
  const esc2 = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const m = rowsToMatrix(rows);
  const head = `<tr>${m[0].map(h => `<th>${esc2(h)}</th>`).join("")}</tr>`;
  const body = m.slice(1).map(r => `<tr>${r.map(v => `<td>${esc2(v)}</td>`).join("")}</tr>`).join("");
  const html = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head>
<meta charset="utf-8">
<style>th{background:#eaeef2;font-weight:700;border:1px solid #9aa;}td{border:1px solid #ccc;}</style>
</head><body><table>${head}${body}</table></body></html>`;
  saveBlob(new Blob(["\uFEFF" + html], { type: "application/vnd.ms-excel;charset=utf-8;" }),
           base + ".xls");
}
function copyRows(rows) {
  const tsv = rowsToMatrix(rows).map(r => r.join("\t")).join("\n");
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsv).catch(() => fallbackCopy(tsv));
  } else fallbackCopy(tsv);
}
function fallbackCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
  document.body.appendChild(ta); ta.select();
  try { document.execCommand("copy"); } catch (e) { /* ignore */ }
  ta.remove();
}
function lotTable(lots, product) {
  const tbl = document.createElement("table"); tbl.className = "tbl lot-tbl";
  tbl.innerHTML = `<thead><tr>
    <th><input type="checkbox" class="hdr-check" title="Toggle all in this tab"></th>
    <th>Lot</th><th>Product ID</th>
    <th class="num">EIMS</th><th class="num">Last used (d)</th>
    <th>Operation</th><th>PI Dispose</th></tr></thead>`;
  const tb = document.createElement("tbody");
  lots.forEach(r => {
    const tr = document.createElement("tr");
    const mis = r.status === "MISMATCH";
    if (mis) tr.className = "row-mis";
    const checked = STATE.picks[r.lot] ? "checked" : "";
    // The RUPS count is no longer its own column; when it disagrees with EIMS
    // it is shown inline next to the EIMS quantity.
    const qty = mis
      ? `${r.eims_qty == null ? "" : r.eims_qty}<span class="qty-rups" title="RUPS unit count">RUPS ${r.rups_qty}</span>`
      : (r.eims_qty == null ? "" : r.eims_qty);
    tr.innerHTML = `<td><input type="checkbox" ${checked} data-lot="${esc(r.lot)}"></td>
      <td>${esc(r.lot)}</td>
      <td class="pid">${esc(r.product_id || "")}</td>
      <td class="num">${qty}</td>
      <td class="num">${r.last_used == null ? "" : r.last_used}</td>
      <td>${esc(r.operation)}</td>
      <td>${esc(r.pi_dispose)}</td>`;
    if (mis) tr.title = `Quantity mismatch — EIMS ${r.eims_qty}, RUPS ${r.rups_qty}`;
    else if (r.status === "Not found") tr.title = "Lot not found in RUPS";
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
  // Header checkbox toggles every row in this tab in ONE batch.
  // (Calling applyCheck per row re-entered refreshSelection, which reset
  //  hdr.checked mid-loop and made select-all behave like single ticks.)
  const hdr = tbl.querySelector(".hdr-check");
  const toggleAll = want => {
    $$("tbody input[type=checkbox]", tbl).forEach(c => {
      c.checked = want;
      const rec = c._rec;
      if (!rec) return;
      if (want) STATE.picks[rec.lot] = rec; else delete STATE.picks[rec.lot];
    });
    _lastCb = null;
    refreshSelection();
  };
  hdr.onclick = ev => ev.stopPropagation();
  hdr.onchange = () => toggleAll(hdr.checked);
  // The checkbox is small, so the whole header cell is a hit target.
  const hcell = hdr.closest("th");
  if (hcell) hcell.onclick = ev => {
    if (ev.target === hdr) return;          // the checkbox handles itself
    toggleAll(!(hdr.checked && !hdr.indeterminate));
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
    cb.checked = want;
    if (want) STATE.picks[rec.lot] = rec; else delete STATE.picks[rec.lot];
  });
  refreshSelection();
}

function refreshSelection() {
  const n = Object.keys(STATE.picks).length;
  const c = $("#selCount"); if (c) c.textContent = n + " selected";
  // Sync header checkboxes, including the partial (indeterminate) state.
  $$("#reconProducts table.lot-tbl").forEach(tbl => {
    const hdr = tbl.querySelector(".hdr-check");
    if (!hdr) return;
    const boxes = $$("tbody input[type=checkbox]", tbl);
    const on = boxes.filter(b => b.checked).length;
    hdr.checked = boxes.length > 0 && on === boxes.length;
    hdr.indeterminate = on > 0 && on < boxes.length;
    hdr.title = hdr.checked ? "Clear all rows in this tab"
                            : "Select all rows in this tab";
  });
  // Keep the "exports N rows" hints in sync with the selection.
  $$("#reconProducts .export-bar").forEach(b => { if (b._sync) b._sync(); });
  updateActionBar(n);
  updateAtmfPreview();
}

/* ---------- sticky action bar + ATMf drawer ---------- */
function updateActionBar(n) {
  const bar = $("#actionBar"); if (!bar) return;
  bar.classList.toggle("hidden", n === 0);
  if (!n) { if (isDrawerOpen()) closeAtmf(); return; }
  $("#abCount").textContent = n;
  const byProduct = {};
  Object.values(STATE.picks).forEach(p => {
    byProduct[p.product] = (byProduct[p.product] || 0) + 1;
  });
  const names = Object.entries(byProduct)
    .sort((a, b) => b[1] - a[1])
    .map(([k, v]) => `${k} (${v})`);
  const shown = names.slice(0, 3).join(", ");
  $("#abDetail").textContent =
    names.length > 3 ? `${shown} +${names.length - 3} more` : shown;
}
function isDrawerOpen() {
  const d = $("#atmfCard");
  return d && !d.classList.contains("hidden");
}
function openAtmf() {
  if (!Object.keys(STATE.picks).length) { toast("Select at least one lot first."); return; }
  $("#atmfBackdrop").classList.remove("hidden");
  const d = $("#atmfCard");
  d.classList.remove("hidden"); d.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";
  updateAtmfPreview();
}
function closeAtmf() {
  $("#atmfBackdrop").classList.add("hidden");
  const d = $("#atmfCard");
  d.classList.add("hidden"); d.setAttribute("aria-hidden", "true");
  document.body.style.overflow = "";
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
  const info = $("#atmfFootInfo");
  if (info) {
    const nLots = Object.keys(picks).length;
    info.textContent = list.length
      ? `${list.length} ticket(s) from ${nLots} lot(s)` +
        (allMapped ? "" : " — pick an ATMf Product Name for every row")
      : "No lots selected.";
  }
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
  $$("#reconToolbar button[data-sel]").forEach(b =>
    b.onclick = () => selectLots(b.dataset.sel));

  // Product picker: live search + bulk actions.
  $("#prodSearch").oninput = e => applyProductSearch(e.target.value);
  $("#prodSearch").onkeydown = e => {
    if (e.key === "Escape") { e.target.value = ""; applyProductSearch(""); }
  };
  $$(".pp-bar button[data-pp]").forEach(b => b.onclick = () => {
    const m = b.dataset.pp;
    if (m === "all") setProducts(STATE.allProducts);
    else if (m === "none") setProducts([]);
    else setProducts(STATE.defaultProducts || []);
  });

  // Sticky action bar + ATMf drawer.
  $("#abSubmit").onclick = openAtmf;
  $("#abClear").onclick = () => selectLots("clear");
  $("#atmfClose").onclick = closeAtmf;
  $("#atmfBackdrop").onclick = closeAtmf;
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && isDrawerOpen()) closeAtmf();
  });

  if (window.HAS_EIMS) {
    await loadOptions(); await loadSummary(); await runReconcile();
  } else {
    toast("No EIMS file yet — click Refresh EIMS.");
  }
}
document.addEventListener("DOMContentLoaded", init);
