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
  picks: {},           // lot -> {product, pi, src} ticked for ATMf
  searchProducts: [],  // products from the last "search by lot"
  submitted: {},       // lot -> ticket info, for lots already sent to ATMf
  submitting: false,   // true while an ATMf submit is in flight
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

/* ---------- lot search ---------- */
/* Renders with the very same product -> PI -> lots component as the
   reconciliation view, so searched lots stay tickable and submittable. */
async function runLotSearch() {
  const q = $("#lotQuery").value.trim();
  if (!q) { toast("Enter at least one lot number."); return; }
  overlay(true, "Searching\u2026");
  try {
    const data = await api("/api/lot/search", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ query: q }),
    });
    renderLotSearch(data);
  } catch (e) { toast("Search failed: " + e.message); }
  finally { overlay(false); }
}

function clearLotSearch() {
  $("#lotQuery").value = "";
  $("#searchHint").textContent = "";
  $("#searchProducts").innerHTML = "";
  $("#searchResults").classList.add("hidden");
  STATE.searchProducts = [];
  // Drop any picks that came from the search so the action bar stays honest.
  Object.entries(STATE.picks).forEach(([lot, r]) => { if (r.src === "search") delete STATE.picks[lot]; });
  refreshSelection();
}

function renderLotSearch(d) {
  const prods = d.products || [];
  const hint = $("#searchHint");
  const notes = [`${d.lots || 0} lot(s) from ${(d.query || []).length} term(s)`];
  Object.entries(d.expanded || {}).forEach(([tok, n]) => notes.push(`"${tok}" \u2192 ${n}`));
  if (d.reused) notes.push(`${d.reused} from cache`);
  if (d.fetched) notes.push(`${d.fetched} newly reconciled`);
  if (d.not_found && d.not_found.length) notes.push(`${d.not_found.length} not found`);
  if (d.truncated) notes.push("truncated \u2014 narrow the search");
  hint.textContent = notes.join(" \u00b7 ");

  const wrap = $("#searchProducts"); wrap.innerHTML = "";
  $("#searchResults").classList.remove("hidden");

  // "Did you mean" chips for terms that matched nothing.
  const sugg = d.suggestions || {};
  if (Object.keys(sugg).length) {
    const box = document.createElement("div"); box.className = "sugg-box";
    box.innerHTML = "<b>Did you mean:</b>";
    Object.entries(sugg).forEach(([tok, near]) => near.forEach(n => {
      const b = document.createElement("button");
      b.className = "chip"; b.textContent = n;
      b.onclick = () => { $("#lotQuery").value = n; runLotSearch(); };
      box.appendChild(b);
    }));
    wrap.appendChild(box);
  }

  if (!prods.length) {
    wrap.insertAdjacentHTML("beforeend",
      `<p class="muted">${esc(d.message || "No lots found.")}</p>`);
    return;
  }

  STATE.searchProducts = prods;
  prods.forEach(p => wrap.appendChild(productBlock(p, "search")));
  refreshSelection();
}

/* One product block: header + PI tabs + lot table + export bar.
   src marks where a pick came from ("recon" or "search"). */
function productBlock(p, src) {
  const box = document.createElement("div"); box.className = "prod-detail";
  const head = document.createElement("div"); head.className = "prod-head";
  head.innerHTML = `<span class="prod-name">\u{1F4E6} ${esc(p.product)}</span>
    <span class="muted">${p.n_lots} lot(s)</span>
    <span class="badge ppv">PPV ${p.n_ppv}</span>
    <span class="badge cls">Class ${p.n_class}</span>
    ${p.n_eng ? `<span class="badge eng">Eng ${p.n_eng}</span>` : ""}`;
  const selBtn = document.createElement("button");
  selBtn.className = "btn btn-sm"; selBtn.textContent = "Select all in this product";
  selBtn.onclick = () => selectLots("all", box);
  head.appendChild(selBtn);
  box.appendChild(head);

  const tabs = document.createElement("div"); tabs.className = "pi-tabs";
  const panes = document.createElement("div");
  const pending = [];
  p.groups.forEach((g, i) => {
    const tab = document.createElement("div");
    tab.className = "pi-tab" + (i === 0 ? " active" : "");
    tab.textContent = `${g.pi} (${g.lots.length})`;
    const pn = document.createElement("div");
    pn.style.display = i === 0 ? "block" : "none";
    const scroll = document.createElement("div"); scroll.className = "lot-scroll";
    const table = lotTable(g.lots, p.product, src);
    scroll.appendChild(table);
    pn.appendChild(scroll);
    pn.appendChild(exportBar(g.lots, p.product, g.pi));
    pending.push(table);
    tab.onclick = () => {
      $$(".pi-tab", tabs).forEach(t => t.classList.remove("active"));
      $$("div", panes).forEach(x => { if (x.parentElement === panes) x.style.display = "none"; });
      tab.classList.add("active"); pn.style.display = "block";
    };
    tabs.appendChild(tab); panes.appendChild(pn);
  });
  box.appendChild(tabs); box.appendChild(panes);
  // The toolbar is inserted next to the table, so wire it once the nodes have
  // a parent (the caller appends `box` right after this returns).
  setTimeout(() => pending.forEach(t => enhanceTable(t, { placeholder: "Search lots\u2026" })), 0);
  return box;
}

/* ---------- Excel-like table tools: search, sort, per-column filters ----------
   enhanceTable(table) adds a toolbar (quick search + row counter + reset) and
   makes every header sortable with an Excel-style filter dropdown.
   Columns holding a checkbox are skipped automatically.                       */

function cellText(tr, i) {
  const td = tr.children[i];
  return td ? (td.innerText || td.textContent || "").trim() : "";
}
function cellNum(s) {
  const t = String(s).replace(/[^0-9.\-]/g, "");
  if (!t || t === "-" || t === ".") return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}
// A column sorts numerically only when every non-empty value looks like a number.
function columnIsNumeric(rows, i) {
  let seen = 0;
  for (const tr of rows) {
    const v = cellText(tr, i);
    if (!v) continue;
    if (cellNum(v) === null) return false;
    seen++;
  }
  return seen > 0;
}

let _openMenu = null;
function closeColumnMenu() {
  if (_openMenu) { _openMenu.remove(); _openMenu = null; }
}
document.addEventListener("click", ev => {
  if (_openMenu && !_openMenu.contains(ev.target) &&
      !ev.target.closest(".th-filter")) closeColumnMenu();
});

function enhanceTable(tbl, opts) {
  opts = opts || {};
  const thead = tbl.tHead, tbody = tbl.tBodies[0];
  if (!thead || !thead.rows.length || !tbody) return null;

  // Re-runnable: tables such as the summary are re-rendered on every filter
  // change, so drop the previous toolbar before wiring a fresh one.
  if (tbl._toolbar) { tbl._toolbar.remove(); tbl._toolbar = null; }

  const ths = [...thead.rows[0].cells];
  const skip = new Set(opts.skipCols || []);
  ths.forEach((th, i) => {
    // Selection columns and anything explicitly opted out stay plain.
    if (th.querySelector("input[type=checkbox]") || th.classList.contains("no-tools")) skip.add(i);
  });

  let rows = [...tbody.rows];
  const state = { q: "", sortCol: -1, dir: 0, filters: {} };

  // ---- toolbar -------------------------------------------------------------
  // Sits above the scroll box when the table is inside one.
  const host = tbl.closest(".lot-scroll") || tbl.closest(".table-scroll") || tbl;
  const bar = document.createElement("div"); bar.className = "tt-bar";
  const search = document.createElement("input");
  search.type = "search"; search.className = "tt-search";
  search.placeholder = opts.placeholder || "Search this table\u2026";
  const count = document.createElement("span"); count.className = "tt-count";
  const reset = document.createElement("button");
  reset.className = "btn btn-sm btn-ghost tt-reset hidden"; reset.textContent = "Reset";
  reset.title = "Clear the search, sorting and every column filter";
  bar.append(search, count, reset);
  host.parentNode.insertBefore(bar, host);
  tbl._toolbar = bar;

  search.oninput = () => { state.q = search.value.trim().toLowerCase(); apply(); };
  search.onkeydown = e => { if (e.key === "Escape") { search.value = ""; state.q = ""; apply(); } };
  reset.onclick = () => {
    state.q = ""; state.sortCol = -1; state.dir = 0; state.filters = {};
    search.value = ""; apply();
  };

  // ---- headers -------------------------------------------------------------
  ths.forEach((th, i) => {
    if (skip.has(i)) return;
    // Keep the original label when this header was already decorated.
    const existing = th.querySelector(".th-label");
    const label = existing ? existing.innerHTML : th.innerHTML;
    th.classList.add("th-tool");
    th.innerHTML = `<span class="th-label">${label}</span>` +
      `<span class="th-sort"></span>` +
      `<button class="th-filter" title="Filter this column">\u25be</button>`;
    th.querySelector(".th-label").onclick = () => {
      // click cycles: ascending -> descending -> unsorted
      if (state.sortCol !== i) { state.sortCol = i; state.dir = 1; }
      else if (state.dir === 1) state.dir = -1;
      else { state.sortCol = -1; state.dir = 0; }
      apply();
    };
    th.querySelector(".th-filter").onclick = ev => {
      ev.stopPropagation();
      openColumnMenu(th, i);
    };
  });

  // ---- per-column filter menu ---------------------------------------------
  function openColumnMenu(th, i) {
    const wasMine = _openMenu && _openMenu._col === i;
    closeColumnMenu();
    if (wasMine) return;

    const menu = document.createElement("div");
    menu.className = "col-menu"; menu._col = i;
    // Distinct values from rows that pass every OTHER filter, like Excel.
    const values = [...new Set(rows
      .filter(tr => passesFilters(tr, i) && passesSearch(tr))
      .map(tr => cellText(tr, i)))]
      .sort((a, b) => {
        const na = cellNum(a), nb = cellNum(b);
        if (na !== null && nb !== null) return na - nb;
        return String(a).localeCompare(String(b));
      });
    const allowed = state.filters[i];

    menu.innerHTML = `
      <div class="cm-sort">
        <button data-s="asc">\u2191 Sort A \u2192 Z</button>
        <button data-s="desc">\u2193 Sort Z \u2192 A</button>
      </div>
      <input type="search" class="cm-search" placeholder="Search values\u2026">
      <label class="cm-all"><input type="checkbox" class="cm-all-box"> <b>(Select all)</b></label>
      <div class="cm-list"></div>
      <div class="cm-foot">
        <button class="btn btn-sm btn-ghost" data-a="clear">Clear filter</button>
        <button class="btn btn-sm btn-primary" data-a="ok">Apply</button>
      </div>`;

    const list = menu.querySelector(".cm-list");
    values.forEach(v => {
      const on = !allowed || allowed.has(v);
      const lab = document.createElement("label");
      lab.innerHTML = `<input type="checkbox" ${on ? "checked" : ""}> <span>${esc(v || "(blank)")}</span>`;
      lab.querySelector("input")._v = v;
      list.appendChild(lab);
    });

    const boxes = () => $$("input[type=checkbox]", list);
    const syncAll = () => {
      const vis = boxes().filter(b => !b.closest("label").classList.contains("hide"));
      const on = vis.filter(b => b.checked).length;
      const all = menu.querySelector(".cm-all-box");
      all.checked = vis.length > 0 && on === vis.length;
      all.indeterminate = on > 0 && on < vis.length;
    };
    syncAll();
    list.onchange = syncAll;

    menu.querySelector(".cm-all-box").onchange = e => {
      boxes().forEach(b => {
        if (!b.closest("label").classList.contains("hide")) b.checked = e.target.checked;
      });
      syncAll();
    };
    menu.querySelector(".cm-search").oninput = e => {
      const t = e.target.value.trim().toLowerCase();
      boxes().forEach(b => b.closest("label")
        .classList.toggle("hide", !!t && !String(b._v).toLowerCase().includes(t)));
      syncAll();
    };
    $$(".cm-sort button", menu).forEach(b => b.onclick = () => {
      state.sortCol = i; state.dir = b.dataset.s === "asc" ? 1 : -1;
      closeColumnMenu(); apply();
    });
    menu.querySelector('[data-a="clear"]').onclick = () => {
      delete state.filters[i]; closeColumnMenu(); apply();
    };
    menu.querySelector('[data-a="ok"]').onclick = () => {
      const on = boxes().filter(b => b.checked).map(b => b._v);
      // Everything ticked means "no filter" - keeps the funnel icon honest.
      if (on.length === values.length) delete state.filters[i];
      else state.filters[i] = new Set(on);
      closeColumnMenu(); apply();
    };

    document.body.appendChild(menu);
    const r = th.getBoundingClientRect();
    menu.style.top = `${Math.round(r.bottom + window.scrollY + 4)}px`;
    const left = Math.min(r.left + window.scrollX,
                          window.scrollX + document.documentElement.clientWidth - menu.offsetWidth - 12);
    menu.style.left = `${Math.round(Math.max(8, left))}px`;
    _openMenu = menu;
    menu.querySelector(".cm-search").focus();
  }

  // ---- filtering + sorting -------------------------------------------------
  function passesSearch(tr) {
    return !state.q || (tr.innerText || "").toLowerCase().includes(state.q);
  }
  function passesFilters(tr, except) {
    for (const k in state.filters) {
      const i = +k;
      if (i === except) continue;
      if (!state.filters[i].has(cellText(tr, i))) return false;
    }
    return true;
  }

  function apply() {
    let out = rows.filter(tr => passesSearch(tr) && passesFilters(tr, -1));

    if (state.sortCol >= 0 && state.dir) {
      const i = state.sortCol, numeric = columnIsNumeric(rows, i);
      out = out.slice().sort((a, b) => {
        const va = cellText(a, i), vb = cellText(b, i);
        // Blanks always sink to the bottom, whichever direction we sort.
        if (!va && !vb) return 0;
        if (!va) return 1;
        if (!vb) return -1;
        const r = numeric ? (cellNum(va) - cellNum(vb)) : va.localeCompare(vb, undefined, { numeric: true });
        return r * state.dir;
      });
    }

    const keep = new Set(out);
    rows.forEach(tr => { if (!keep.has(tr)) tr.remove(); });
    const frag = document.createDocumentFragment();
    out.forEach(tr => frag.appendChild(tr));
    tbody.appendChild(frag);

    // Header affordances.
    ths.forEach((th, i) => {
      if (skip.has(i)) return;
      const s = th.querySelector(".th-sort");
      if (s) s.textContent = state.sortCol === i ? (state.dir === 1 ? "\u25b2" : "\u25bc") : "";
      const f = th.querySelector(".th-filter");
      if (f) f.classList.toggle("on", !!state.filters[i]);
    });

    const active = state.q || state.sortCol >= 0 || Object.keys(state.filters).length;
    reset.classList.toggle("hidden", !active);
    count.textContent = out.length === rows.length
      ? `${rows.length} row(s)` : `${out.length} of ${rows.length} row(s)`;
    count.classList.toggle("filtered", out.length !== rows.length);

    if (typeof opts.onApply === "function") opts.onApply(out);
    // Only the header tick needs refreshing here; a full refreshSelection()
    // would re-hit the ATMf guess API on every sort or keystroke.
    syncHeaderCheck(tbl);
  }

  tbl._tools = {
    apply,
    // Visible rows only - "select all" should follow what you can see.
    visibleRows: () => [...tbody.rows],
  };
  apply();
  return tbl._tools;
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
  enhanceTable($("#prodSummary"), { placeholder: "Search products\u2026" });
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
  // Same component the search results use.
  const block = productBlock(p, "recon");
  while (block.firstChild) detail.appendChild(block.firstChild);
}

/* ---------- export a lot table ---------- */
const EXPORT_COLS = [
  ["lot", "Lot"], ["product", "Product"], ["product_id", "Product ID"],
  ["eims_qty", "EIMS_Qty"], ["rups_qty", "RUPS_Qty"], ["last_used", "Last_Used_Days"],
  ["operation", "Operation"], ["pi_dispose", "PI_Dispose"], ["status", "Status"],
  ["_ticket", "ATMf_Ticket"],
];

function exportBar(lots, product, pi) {
  const bar = document.createElement("div"); bar.className = "export-bar";
  const info = document.createElement("span"); info.className = "muted";
  bar.appendChild(info);

  // Rows currently shown by the table tools (search / column filters), so an
  // export mirrors exactly what is on screen - like copying a filtered range.
  const visibleLots = () => {
    const tbl = bar.parentElement && bar.parentElement.querySelector("table.lot-tbl");
    if (!tbl) return lots;
    const shown = new Set($$("tbody input[type=checkbox]", tbl)
      .map(cb => cb._rec && cb._rec.lot).filter(Boolean));
    return shown.size ? lots.filter(r => shown.has(r.lot)) : lots;
  };
  const target = () => {
    const vis = visibleLots();
    const picked = vis.filter(r => STATE.picks[r.lot]);
    return { rows: picked.length ? picked : vis, picked: picked.length,
             filtered: vis.length !== lots.length };
  };

  const mk = (label, title, fn) => {
    const b = document.createElement("button");
    b.className = "btn btn-sm btn-ghost"; b.textContent = label; b.title = title;
    b.onclick = () => {
      const t = target();
      fn(t.rows, `${slugName(product)}_${slugName(pi)}_${stamp()}`);
      toast(`Exported ${t.rows.length} row(s)${t.picked ? " (selected only)" : ""}.`);
    };
    return b;
  };
  bar.appendChild(mk("⬇ CSV", "Download what is shown as CSV (opens in Excel)", downloadCsvCols(EXPORT_COLS)));
  bar.appendChild(mk("⬇ Excel", "Download what is shown as an Excel file", downloadXlsCols(EXPORT_COLS)));
  bar.appendChild(mk("⧉ Copy", "Copy what is shown to the clipboard", copyRowsCols(EXPORT_COLS)));

  // Keep the hint in sync with the selection and any active column filter.
  const sync = () => {
    const t = target();
    info.textContent = t.picked
      ? `Exports the ${t.picked} selected row(s)`
      : `Exports ${t.filtered ? "the " + t.rows.length + " filtered" : "all " + t.rows.length} row(s)`;
  };
  sync();
  bar._sync = sync;
  return bar;
}

function rowsToMatrix(rows, cols) {
  const head = cols.map(([, h]) => h);
  const body = rows.map(r => cols.map(([k]) => {
    // Ticket id is held in STATE, not on the row itself.
    if (k === "_ticket") {
      const d = STATE.submitted[r.lot];
      return d ? d.ticket_id : "";
    }
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

// The three exporters are column-agnostic; *Cols() binds a column list so the
// same code serves the reconciliation tables, the lot summary and unit lists.
function downloadCsvCols(cols) {
  return (rows, base) => {
    const csv = rowsToMatrix(rows, cols).map(r =>
      r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\r\n");
    // UTF-8 BOM so Excel opens it with the right encoding.
    saveBlob(new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" }), base + ".csv");
  };
}
function downloadXlsCols(cols) {
  return (rows, base) => {
    // SpreadsheetML table - Excel opens this natively, no library needed.
    const esc2 = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const m = rowsToMatrix(rows, cols);
    const head = `<tr>${m[0].map(h => `<th>${esc2(h)}</th>`).join("")}</tr>`;
    const body = m.slice(1).map(r => `<tr>${r.map(v => `<td>${esc2(v)}</td>`).join("")}</tr>`).join("");
    const html = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head>
<meta charset="utf-8">
<style>th{background:#eaeef2;font-weight:700;border:1px solid #9aa;}td{border:1px solid #ccc;}</style>
</head><body><table>${head}${body}</table></body></html>`;
    saveBlob(new Blob(["\uFEFF" + html], { type: "application/vnd.ms-excel;charset=utf-8;" }),
             base + ".xls");
  };
}
function copyRowsCols(cols) {
  return rows => {
    const tsv = rowsToMatrix(rows, cols).map(r => r.join("\t")).join("\n");
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).catch(() => fallbackCopy(tsv));
    } else fallbackCopy(tsv);
  };
}

function fallbackCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
  document.body.appendChild(ta); ta.select();
  try { document.execCommand("copy"); } catch (e) { /* ignore */ }
  ta.remove();
}
function lotTable(lots, product, src) {
  const tbl = document.createElement("table"); tbl.className = "tbl lot-tbl";
  tbl.innerHTML = `<thead><tr>
    <th><input type="checkbox" class="hdr-check" title="Toggle all in this tab"></th>
    <th>Lot</th><th>Product ID</th>
    <th class="num">EIMS</th><th class="num">Last used (d)</th>
    <th>Operation</th><th>PI Dispose</th><th>Ticket</th></tr></thead>`;
  const tb = document.createElement("tbody");
  lots.forEach(r => {
    const tr = document.createElement("tr");
    const mis = r.status === "MISMATCH";
    if (mis) tr.className = "row-mis";
    // Already sent to ATMf? Flag it so nobody submits the same lot twice.
    const done = STATE.submitted[r.lot];
    if (done) tr.classList.add("row-done");
    const checked = STATE.picks[r.lot] ? "checked" : "";
    // The RUPS count is no longer its own column; when it disagrees with EIMS
    // it is shown inline next to the EIMS quantity.
    const qty = mis
      ? `${r.eims_qty == null ? "" : r.eims_qty}<span class="qty-rups" title="RUPS unit count">RUPS ${r.rups_qty}</span>`
      : (r.eims_qty == null ? "" : r.eims_qty);
    const ticket = done
      ? `<a class="tkt" href="${esc(done.url)}" target="_blank" rel="noopener"
           title="Submitted ${esc(done.at || "")} — ${esc(done.request || "")}">✓ ${esc(done.ticket_id)}</a>`
      : "";
    tr.innerHTML = `<td><input type="checkbox" ${checked} data-lot="${esc(r.lot)}"></td>
      <td>${esc(r.lot)}</td>
      <td class="pid">${esc(r.product_id || "")}</td>
      <td class="num">${qty}</td>
      <td class="num">${r.last_used == null ? "" : r.last_used}</td>
      <td>${esc(r.operation)}</td>
      <td>${esc(r.pi_dispose)}</td>
      <td class="tkt-cell">${ticket}</td>`;
    if (mis) tr.title = `Quantity mismatch — EIMS ${r.eims_qty}, RUPS ${r.rups_qty}`;
    else if (r.status === "Not found") tr.title = "Lot not found in RUPS";
    const cb = tr.querySelector("input");
    cb._rec = { lot: r.lot, product, pi: r.pi_dispose, src: src || "recon" };
    cb.onchange = ev => applyCheck(cb, ev);
    // clicking anywhere on the row toggles the checkbox
    tr.onclick = ev => {
      if (ev.target.tagName === "INPUT" || ev.target.tagName === "A") return;
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

// Everywhere lot tables can live: the reconciliation pane and the search results.
const LOT_AREAS = "#reconProducts, #searchProducts";

function selectLots(mode, root) {
  if (!root) {
    // Global toolbar: operate on ALL reconciliation products (only one is
    // rendered in the DOM), so drive selection from the data, then sync any
    // visible checkboxes.
    (STATE.reconProducts || []).forEach(p =>
      p.groups.forEach(g => g.lots.forEach(r => {
        let want;
        if (mode === "all") want = true;
        else if (mode === "clear") want = false;
        if (want) STATE.picks[r.lot] = { lot: r.lot, product: p.product, pi: r.pi_dispose, src: "recon" };
        else delete STATE.picks[r.lot];
      })));
    if (mode === "clear") STATE.picks = {};
    $$(LOT_AREAS).forEach(area =>
      $$("tbody input[type=checkbox]", area).forEach(cb => {
        const rec = cb._rec; if (rec) cb.checked = STATE.picks[rec.lot] != null;
      }));
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

// Header tick reflects only the rows you can actually see, so it stays
// meaningful while a column filter or search is narrowing the table.
function syncHeaderCheck(tbl) {
  const hdr = tbl.querySelector(".hdr-check");
  if (!hdr) return;
  const boxes = $$("tbody input[type=checkbox]", tbl);
  const on = boxes.filter(b => b.checked).length;
  hdr.checked = boxes.length > 0 && on === boxes.length;
  hdr.indeterminate = on > 0 && on < boxes.length;
  hdr.title = hdr.checked ? "Clear all visible rows" : "Select all visible rows";
}

function refreshSelection() {
  const n = Object.keys(STATE.picks).length;
  const c = $("#selCount"); if (c) c.textContent = n + " selected";
  $$(LOT_AREAS).forEach(area => {
    $$("table.lot-tbl", area).forEach(syncHeaderCheck);
    // Keep the "exports N rows" hints in sync with the selection.
    $$(".export-bar", area).forEach(b => { if (b._sync) b._sync(); });
  });
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
  // Never tear the drawer down mid-request.
  if (STATE.submitting) { toast("Still submitting\u2026 please wait."); return; }
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
  const tb = $("#atmfPreview tbody"); tb.innerHTML = "";
  const prodMap = STATE._guessMap || {};
  list.forEach((g, i) => {
    const name = prodMap[g.product] || "";
    const dupes = g.lots.filter(l => STATE.submitted[l]);
    const tr = document.createElement("tr");
    if (dupes.length) tr.className = "row-done";
    const lotCell = g.lots.map(l => STATE.submitted[l]
      ? `<span class="lot-done" title="Already on ticket ${esc(STATE.submitted[l].ticket_id)}">${esc(l)}</span>`
      : esc(l)).join(", ");
    tr.innerHTML = `<td>${reqSelect(i)}</td><td>${esc(g.product)}</td>
      <td>${nameInput(i, name)}</td><td>${esc(g.area)}</td>
      <td>${lotCell}</td><td class="num">${g.lots.length}</td>`;
    tb.appendChild(tr);
    g._name = name;
    g._dupes = dupes;
  });
  STATE._atmfGroups = list;
  renderDupeWarning(list);

  // Re-check the submit button whenever a product name is edited, otherwise a
  // manual pick would never enable it.
  $$('#atmfPreview input[data-name]').forEach(inp => {
    inp.oninput = () => validateAtmf();
    inp.onchange = () => validateAtmf();
  });

  // Only worth a toolbar once there is something to sift through.
  if (list.length > 3) enhanceTable($("#atmfPreview"), { placeholder: "Search tickets\u2026" });
  else { const t = $("#atmfPreview"); if (t && t._toolbar) { t._toolbar.remove(); t._toolbar = null; } }
  validateAtmf();
}

/* Warn before creating a second ticket for lots that already have one. */
function renderDupeWarning(list) {
  const box = $("#atmfDupes");
  if (!box) return;
  const dupes = [...new Set(list.flatMap(g => g._dupes || []))];
  if (!dupes.length) { box.classList.add("hidden"); box.innerHTML = ""; return; }
  box.classList.remove("hidden");
  box.innerHTML =
    `<b>⚠ ${dupes.length} lot(s) already have a ticket.</b>
     <span class="muted">Submitting again creates a duplicate.</span>
     <div class="dupe-list">${dupes.map(l => {
       const d = STATE.submitted[l];
       return `<a href="${esc(d.url)}" target="_blank" rel="noopener"
                 title="Submitted ${esc(d.at || "")}">${esc(l)} → ${esc(d.ticket_id)}</a>`;
     }).join("")}</div>
     <div class="dupe-actions">
       <button class="btn btn-sm btn-ghost" data-d="drop">Remove them from this batch</button>
       <button class="btn btn-sm btn-ghost" data-d="forget">They were cancelled — clear the flag</button>
     </div>`;
  box.querySelector('[data-d="drop"]').onclick = () => {
    dupes.forEach(l => delete STATE.picks[l]);
    $$(LOT_AREAS).forEach(area =>
      $$("tbody input[type=checkbox]", area).forEach(cb => {
        const rec = cb._rec; if (rec) cb.checked = STATE.picks[rec.lot] != null;
      }));
    refreshSelection();
    toast(`Removed ${dupes.length} already-submitted lot(s).`);
  };
  box.querySelector('[data-d="forget"]').onclick = async () => {
    try {
      await api("/api/atmf/forget", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ lots: dupes }),
      });
      await loadSubmitted();
      updateAtmfPreview();
      toast("Cleared — those lots can be submitted again.");
    } catch (e) { toast("Could not clear: " + e.message); }
  };
}

async function loadSubmitted() {
  try {
    const d = await api("/api/atmf/submitted");
    STATE.submitted = d.submitted || {};
  } catch (e) { STATE.submitted = {}; }
}

/* A row is ready when its product name is one of the real ATMf products.
   Driven by what is in the box right now, not by the original guess. */
function validateAtmf() {
  const valid = new Set(STATE.atmfProducts || []);
  const inputs = $$('#atmfPreview input[data-name]');
  let missing = 0;
  inputs.forEach(inp => {
    const ok = valid.has(inp.value.trim());
    inp.classList.toggle("bad", !ok && inp.value.trim() !== "");
    inp.classList.toggle("empty", inp.value.trim() === "");
    if (!ok) missing++;
  });
  // Lots that already have a ticket block the button outright; clearing them
  // is a deliberate click, which is exactly what stops accidental duplicates.
  const dupes = [...new Set((STATE._atmfGroups || []).flatMap(g => g._dupes || []))];
  const ready = inputs.length > 0 && missing === 0 && dupes.length === 0
                && !STATE.submitting;
  $("#btnSubmitAtmf").disabled = !ready;

  const info = $("#atmfFootInfo");
  if (info) {
    const nLots = Object.keys(STATE.picks).length;
    const nT = (STATE._atmfGroups || []).length;
    let msg;
    if (!nT) msg = "No lots selected.";
    else if (STATE.submitting) msg = "Submitting\u2026 please wait.";
    else {
      msg = `${nT} ticket(s) from ${nLots} lot(s)`;
      if (missing) msg += ` \u2014 ${missing} row(s) still need a valid ATMf Product Name`;
      else if (dupes.length) msg += ` \u2014 ${dupes.length} lot(s) already submitted, resolve them above`;
    }
    info.textContent = msg;
  }
  return ready;
}

function reqSelect(i) {
  return `<select data-req="${i}">${INTMS.requests.map(r => `<option>${esc(r)}</option>`).join("")}</select>`;
}

/* Type-to-search box backed by a <datalist>: the browser filters 432 product
   names as you type, which a plain <select> cannot do. */
function nameInput(i, sel) {
  return `<input class="atmf-name" data-name="${i}" list="atmfProductList"
    value="${esc(sel)}" placeholder="Type to search\u2026" autocomplete="off"
    spellcheck="false" title="Start typing a product name, e.g. ADL or 8+16">`;
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
  // Guard against a second click landing before the first request returns.
  // (The busy overlay alone is not enough - the drawer sits above it.)
  if (STATE.submitting) return;
  if (!validateAtmf()) { toast("Every row needs a valid ATMf Product Name."); return; }
  const list = STATE._atmfGroups || [];
  const tickets = list.map((g, i) => {
    const reqSel = $(`#atmfPreview select[data-req="${i}"]`);
    const nameInp = $(`#atmfPreview input[data-name="${i}"]`);
    return {
      product_name: nameInp ? nameInp.value.trim() : g._name,
      test_area: g.area,
      request: reqSel ? reqSel.value : INTMS.requests[0],
      lots: g.lots,
      product: g.product,   // Prodgroup3 - so the server can learn the mapping
      customer: $("#atmfCustomer").value, payer: $("#atmfPayer").value,
      stage: $("#atmfStage").value, factory: $("#atmfFactory").value,
      shipping: $("#atmfShipping").value, nfo: $("#atmfNfo").value,
    };
  });
  if (tickets.some(t => !t.product_name)) { toast("Pick a Product Name for every row."); return; }

  // Lock the button for the whole round trip. Creating a ticket is not
  // idempotent, so a double click would create a duplicate.
  STATE.submitting = true;
  const btn = $("#btnSubmitAtmf");
  const label = btn.textContent;
  btn.disabled = true;
  btn.classList.add("is-busy");
  btn.textContent = `Submitting ${tickets.length} ticket(s)\u2026`;
  overlay(true, `Creating ${tickets.length} ATMf ticket(s)\u2026`);
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
    const okN = data.results.filter(r => r.ok).length;
    toast(okN === data.results.length
      ? `Created ${okN} ticket(s).`
      : `Created ${okN} of ${data.results.length} ticket(s) — see the results below.`);
    // Pick up the new "already submitted" markers and repaint the lot tables.
    await loadSubmitted();
    if (STATE.reconProducts.length) renderReconProducts(STATE.reconProducts);
    if (STATE.searchProducts.length) {
      const w = $("#searchProducts"); w.innerHTML = "";
      STATE.searchProducts.forEach(p => w.appendChild(productBlock(p, "search")));
    }
    updateAtmfPreview();
  } catch (e) { toast("Submit failed: " + e.message); }
  finally {
    overlay(false);
    STATE.submitting = false;
    btn.classList.remove("is-busy");
    btn.textContent = label;
    // validateAtmf decides whether it may be enabled again; every lot that
    // just went out is now flagged, so a repeat needs a deliberate action.
    validateAtmf();
  }
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
  await loadSubmitted();
  // Fill the datalist once; every Product Name box shares it.
  const dl = $("#atmfProductList");
  if (dl) dl.innerHTML = STATE.atmfProducts
    .map(n => `<option value="${esc(n)}"></option>`).join("");

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

  // Lot search.
  $("#btnLotSearch").onclick = runLotSearch;
  $("#btnLotClear").onclick = clearLotSearch;
  $("#lotQuery").onkeydown = e => { if (e.key === "Enter") runLotSearch(); };

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
