
/**
 * 診断データ解析ツール（静的サイト）
 * - Excel(.xlsx)をブラウザ内で解析（SheetJS）
 * - テーブル探索（DataTables）
 * - プロファイル/可視化（Plotly）
 */

const state = {
  workbook: null,
  sheets: {},       // { sheetName: { headers:[], rows:[{...}] } }
  sheetOrder: [],
  currentSheet: null,
  dataTable: null,

  // Referral deep dive
  referrerTable: null,
  refEdgesTable: null,
  referralDerived: null,           // computed caches for current filter
  referralSelectedReferrer: null,  // selected referrerId
};

function $(id) { return document.getElementById(id); }

function setStatus(msg, kind = "muted") {
  const el = $("loadStatus");
  el.className = "small text-" + kind;
  el.textContent = msg;
}

function isMissing(v) {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "") || (typeof v === "number" && Number.isNaN(v));
}

function toISODateString(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function parseDateInput(v) {
  // "YYYY-MM-DD" -> local midnight Date
  if (!v) return null;
  const m = String(v).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  if (!Number.isFinite(y) || !Number.isFinite(mo) || !Number.isFinite(d)) return null;
  return new Date(y, mo - 1, d);
}

function addDays(d, n) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}

function formatInt(n) {
  if (n === null || n === undefined || Number.isNaN(n)) return "-";
  return Number(n).toLocaleString();
}

function formatPct(p, digits = 1) {
  if (p === null || p === undefined || Number.isNaN(p)) return "-";
  return (Number(p) * 100).toFixed(digits) + "%";
}

function formatHours(h, digits = 2) {
  if (h === null || h === undefined || Number.isNaN(h)) return "-";
  return Number(h).toFixed(digits);
}

function parseJsonSafe(s) {
  if (s === null || s === undefined) return {};
  if (typeof s !== "string") return {};
  const txt = s.trim();
  if (!txt) return {};
  try { return JSON.parse(txt); } catch (e) { return {}; }
}

function normalizeCellValue(v) {
  if (v instanceof Date) return v.toISOString();
  return v;
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

function getUIFlags() {
  return {
    maskPII: $("maskPII").checked,
    showJsonCols: $("showJsonCols").checked,
    showUnnamedCols: $("showUnnamedCols").checked,
  };
}

function shouldHideColumn(colName, flags) {
  const c = String(colName || "");
  if (!flags.showJsonCols) {
    const lower = c.toLowerCase();
    if (lower.includes("raw_json") || lower.endsWith("_json") || lower.includes("payload_json") || lower.includes("answers_json")) {
      return true;
    }
  }
  if (!flags.showUnnamedCols) {
    if (c.startsWith("Unnamed:")) return true;
  }
  return false;
}

function maskEmail(email) {
  const s = String(email);
  const at = s.indexOf("@");
  if (at === -1) return s.length <= 2 ? "**" : (s.slice(0, 2) + "****");
  const local = s.slice(0, at);
  const domain = s.slice(at);
  if (local.length <= 2) return local[0] + "*".repeat(Math.max(1, local.length - 1)) + domain;
  return local.slice(0, 2) + "*".repeat(Math.min(8, Math.max(2, local.length - 2))) + domain;
}

function maskName(name) {
  const s = String(name);
  if (s.length === 0) return s;
  if (s.length === 1) return s + "*";
  return s.slice(0, 1) + "*".repeat(Math.min(6, s.length - 1));
}

function maskIdLike(v) {
  const s = String(v);
  if (s.length <= 10) return s;
  return s.slice(0, 6) + "…" + s.slice(-4);
}

function applyMaskForDisplay(col, v, flags) {
  if (!flags.maskPII) return v;
  const c = String(col || "").toLowerCase();
  if (isMissing(v)) return v;

  // Obvious PII
  if (c === "email" || c.endsWith("email")) return maskEmail(v);
  if (c === "name" || c.endsWith("name")) return maskName(v);

  // IDs (userId/referrerId/gmail ids etc.)
  if (c.includes("gmail_message_id")) return maskIdLike(v);
  if (c === "userid" || c.endsWith("userid")) return maskIdLike(v);
  if (c === "referrerid" || c.endsWith("referrerid")) return maskIdLike(v);
  if (/(^|_)id$/.test(c)) return maskIdLike(v);

  return v;
}

function detectColumnType(values) {
  // values: non-missing sample
  // returns: "number" | "date" | "boolean" | "string"
  let nNum = 0, nDate = 0, nBool = 0, nStr = 0;
  for (const v of values) {
    if (typeof v === "number" && Number.isFinite(v)) { nNum++; continue; }
    if (typeof v === "boolean") { nBool++; continue; }
    if (typeof v === "string") {
      const s = v.trim();
      // numeric
      if (/^-?\d+(\.\d+)?$/.test(s)) { nNum++; continue; }
      // ISO-ish date
      if (/^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{4}\/\d{2}\/\d{2}/.test(s)) {
        const t = Date.parse(s);
        if (!Number.isNaN(t)) { nDate++; continue; }
      }
      // fallback
      nStr++;
      continue;
    }
    // others -> string
    nStr++;
  }
  const total = Math.max(1, values.length);
  if (nNum / total >= 0.8) return "number";
  if (nDate / total >= 0.8) return "date";
  if (nBool / total >= 0.8) return "boolean";
  return "string";
}

function safeNumber(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "string") {
    const s = v.trim();
    if (s === "") return null;
    const x = Number(s);
    return Number.isFinite(x) ? x : null;
  }
  return null;
}

function safeDate(v) {
  if (v instanceof Date) return v;
  if (typeof v === "string") {
    const t = Date.parse(v);
    if (!Number.isNaN(t)) return new Date(t);
  }
  return null;
}

function median(sortedNums) {
  const n = sortedNums.length;
  if (n === 0) return null;
  const mid = Math.floor(n / 2);
  if (n % 2 === 1) return sortedNums[mid];
  return (sortedNums[mid - 1] + sortedNums[mid]) / 2;
}

function quantile(sortedNums, q) {
  const n = sortedNums.length;
  if (n === 0) return null;
  const pos = (n - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (sortedNums[base + 1] === undefined) return sortedNums[base];
  return sortedNums[base] + rest * (sortedNums[base + 1] - sortedNums[base]);
}

function freqTop(values, topN = 5) {
  const m = new Map();
  for (const v of values) {
    const key = String(v);
    m.set(key, (m.get(key) || 0) + 1);
  }
  const arr = Array.from(m.entries()).sort((a, b) => b[1] - a[1]).slice(0, topN);
  return arr.map(([k, c]) => ({ value: k, count: c }));
}

function corrPearson(x, y) {
  // x,y arrays of numbers with same length; may contain nulls
  let n = 0;
  let sx = 0, sy = 0, sxx = 0, syy = 0, sxy = 0;
  for (let i = 0; i < x.length; i++) {
    const xi = x[i], yi = y[i];
    if (xi === null || yi === null) continue;
    n++;
    sx += xi; sy += yi;
    sxx += xi * xi;
    syy += yi * yi;
    sxy += xi * yi;
  }
  if (n < 3) return null;
  const cov = sxy - (sx * sy) / n;
  const vx = sxx - (sx * sx) / n;
  const vy = syy - (sy * sy) / n;
  if (vx <= 0 || vy <= 0) return null;
  return cov / Math.sqrt(vx * vy);
}

function buildCorrelationMatrix(rows, numericCols) {
  const cols = numericCols;
  const series = {};
  for (const c of cols) {
    series[c] = rows.map(r => safeNumber(r[c]));
  }
  const z = [];
  for (let i = 0; i < cols.length; i++) {
    const row = [];
    for (let j = 0; j < cols.length; j++) {
      const v = corrPearson(series[cols[i]], series[cols[j]]);
      row.push(v === null ? null : Math.round(v * 1000) / 1000);
    }
    z.push(row);
  }
  return { cols, z };
}

function toCSV(rows, cols) {
  const esc = (s) => {
    const str = (s === null || s === undefined) ? "" : String(s);
    if (/[",\n]/.test(str)) return `"${str.replace(/"/g, '""')}"`;
    return str;
  };
  const header = cols.map(esc).join(",");
  const lines = rows.map(r => cols.map(c => esc(r[c])).join(","));
  return [header, ...lines].join("\n");
}

function downloadText(filename, text, mime = "text/plain") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function getVisibleHeaders(sheetName, flags) {
  const headers = state.sheets[sheetName]?.headers || [];
  return headers.filter(h => !shouldHideColumn(h, flags));
}

function getDisplayRows(sheetName, flags) {
  const rawRows = state.sheets[sheetName]?.rows || [];
  const headers = getVisibleHeaders(sheetName, flags);
  return rawRows.map(r => {
    const o = {};
    for (const h of headers) {
      o[h] = applyMaskForDisplay(h, r[h], flags);
    }
    return o;
  });
}

function computeSheetMissingRate(sheetName) {
  const flags = getUIFlags();
  const headers = getVisibleHeaders(sheetName, flags);
  const rows = state.sheets[sheetName]?.rows || [];
  if (rows.length === 0 || headers.length === 0) return 0;
  let missing = 0;
  const total = rows.length * headers.length;
  for (const r of rows) {
    for (const h of headers) {
      if (isMissing(r[h])) missing++;
    }
  }
  return missing / total;
}

function renderOverview() {
  $("overviewEmpty").classList.add("d-none");
  $("overviewContent").classList.remove("d-none");

  const sheetCount = state.sheetOrder.length;
  $("kpiSheetCount").textContent = sheetCount;

  let totalRows = 0;
  for (const s of state.sheetOrder) totalRows += (state.sheets[s].rows.length || 0);
  $("kpiTotalRows").textContent = totalRows;

  $("kpiDiagnosisRows").textContent = state.sheets["diagnosis"] ? state.sheets["diagnosis"].rows.length : "-";
  $("kpiReferralEventsRows").textContent = state.sheets["referral_events"] ? state.sheets["referral_events"].rows.length : "-";

  // Sheet summary table
  const tbody = $("sheetSummaryTable").querySelector("tbody");
  tbody.innerHTML = "";
  for (const s of state.sheetOrder) {
    const rows = state.sheets[s].rows.length;
    const cols = state.sheets[s].headers.length;
    const miss = computeSheetMissingRate(s);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><a href="#" class="sheetLink" data-sheet="${s}">${s}</a></td>
      <td class="text-end">${rows}</td>
      <td class="text-end">${cols}</td>
      <td class="text-end">${(miss * 100).toFixed(1)}%</td>
    `;
    tbody.appendChild(tr);
  }
  tbody.querySelectorAll(".sheetLink").forEach(a => {
    a.addEventListener("click", (e) => {
      e.preventDefault();
      const name = e.currentTarget.getAttribute("data-sheet");
      setCurrentSheet(name);
      // switch tab
      const tabEl = document.querySelector('#tab-sheet');
      const tab = new bootstrap.Tab(tabEl);
      tab.show();
    });
  });

  renderInsights();
  renderOverviewPlots();
}

function renderInsights() {
  const el = $("insights");
  const parts = [];

  // diagnosis insights
  if (state.sheets["diagnosis"]) {
    const rows = state.sheets["diagnosis"].rows;
    const headers = state.sheets["diagnosis"].headers;
    const has = (c) => headers.includes(c);
    parts.push(`<div class="mb-2"><span class="badge badge-soft me-1">diagnosis</span> ${rows.length.toLocaleString()}件</div>`);

    if (has("email")) {
      const uniq = new Set(rows.map(r => r["email"]).filter(v => !isMissing(v))).size;
      parts.push(`<div class="mb-1">ユニークemail: <span class="fw-semibold">${uniq.toLocaleString()}</span></div>`);
    }
    if (has("createdAt")) {
      const ds = rows.map(r => safeDate(r["createdAt"])).filter(d => d);
      if (ds.length) {
        ds.sort((a,b)=>a-b);
        parts.push(`<div class="mb-1">期間: ${toISODateString(ds[0])} 〜 ${toISODateString(ds[ds.length-1])}</div>`);
      }
    }
    if (has("interested")) {
      const interested = rows.filter(r => String(r["interested"]) === "1" || r["interested"] === 1 || r["interested"] === true).length;
      parts.push(`<div class="mb-1">interested=1: ${interested.toLocaleString()}件（${(interested/rows.length*100).toFixed(1)}%）</div>`);
    }
    if (has("age")) {
      const ages = rows.map(r => safeNumber(r["age"])).filter(v => v !== null).sort((a,b)=>a-b);
      if (ages.length) {
        const p50 = median(ages);
        parts.push(`<div class="mb-1">年齢（有効${ages.length.toLocaleString()}件）: 平均 ${(ages.reduce((a,b)=>a+b,0)/ages.length).toFixed(1)}, 中央 ${p50.toFixed(0)}</div>`);
      }
    }
    if (has("gender")) {
      const g = rows.map(r => r["gender"]).filter(v => !isMissing(v));
      const top = freqTop(g, 5);
      if (top.length) {
        parts.push(`<div class="mb-1">gender上位: ${top.map(t => `${t.value}(${t.count})`).join(", ")}</div>`);
      }
    }
  }

  // referral insights
  if (state.sheets["referral_events"]) {
    const rows = state.sheets["referral_events"].rows;
    const headers = state.sheets["referral_events"].headers;
    const has = (c) => headers.includes(c);
    parts.push(`<hr class="my-2" />`);
    parts.push(`<div class="mb-2"><span class="badge badge-soft me-1">referral_events</span> ${rows.length.toLocaleString()}件</div>`);
    if (has("eventType")) {
      const top = freqTop(rows.map(r => r["eventType"]).filter(v => !isMissing(v)), 10);
      parts.push(`<div class="mb-1">eventType: ${top.map(t => `${t.value}(${t.count})`).join(", ")}</div>`);
    }
  }

  if (state.sheets["referrer_summary"]) {
    const rows = state.sheets["referrer_summary"].rows;
    const headers = state.sheets["referrer_summary"].headers;
    const has = (c) => headers.includes(c);
    parts.push(`<hr class="my-2" />`);
    parts.push(`<div class="mb-2"><span class="badge badge-soft me-1">referrer_summary</span> ${rows.length.toLocaleString()} referrers</div>`);
    if (has("unique_invited_completes")) {
      const sorted = [...rows].sort((a,b) => (safeNumber(b["unique_invited_completes"])||0) - (safeNumber(a["unique_invited_completes"])||0));
      const top = sorted.slice(0,3).map(r => `${r["referrerId"]}: ${safeNumber(r["unique_invited_completes"])||0}`);
      if (top.length) parts.push(`<div class="mb-1">complete上位: ${top.join(", ")}</div>`);
    }
    if (has("visit_to_complete_rate")) {
      const valid = rows.map(r => safeNumber(r["visit_to_complete_rate"])).filter(v => v !== null);
      if (valid.length) {
        valid.sort((a,b)=>a-b);
        parts.push(`<div class="mb-1">visit→complete率: 中央 ${(median(valid)*100).toFixed(1)}%（有効${valid.length}）</div>`);
      }
    }
  }

  if (!parts.length) {
    el.innerHTML = `<div class="text-muted">代表シートが見つかりません（汎用モードで探索してください）。</div>`;
    return;
  }
  el.innerHTML = parts.join("");
}

function renderOverviewPlots() {
  // diagnosis type top
  if (state.sheets["diagnosis"] && state.sheets["diagnosis"].headers.includes("type")) {
    const rows = state.sheets["diagnosis"].rows;
    const values = rows.map(r => r["type"]).filter(v => !isMissing(v));
    const top = freqTop(values, 12);
    const x = top.map(t => t.value);
    const y = top.map(t => t.count);
    Plotly.newPlot("plotTypeTop", [{
      type: "bar",
      x, y
    }], {
      margin: {l: 40, r: 10, t: 10, b: 80},
      xaxis: { tickangle: -45 }
    }, {displayModeBar: false, responsive: true});
  } else {
    $("plotTypeTop").innerHTML = `<div class="text-muted small">diagnosis/type が見つかりません</div>`;
  }

  // events daily
  if (state.sheets["events_daily"]) {
    const rows = state.sheets["events_daily"].rows;
    const dateCol = state.sheets["events_daily"].headers.includes("date") ? "date" : state.sheets["events_daily"].headers[0];
    const x = rows.map(r => {
      const d = safeDate(r[dateCol]);
      return d ? toISODateString(d) : String(r[dateCol]);
    });
    const seriesCols = ["share", "referral_visit", "referral_complete"].filter(c => state.sheets["events_daily"].headers.includes(c));
    const traces = seriesCols.map(c => ({
      type: "scatter",
      mode: "lines+markers",
      name: c,
      x,
      y: rows.map(r => safeNumber(r[c]) || 0)
    }));
    Plotly.newPlot("plotEventsDaily", traces, {
      margin: {l: 40, r: 10, t: 10, b: 40},
      legend: {orientation: "h"}
    }, {displayModeBar: false, responsive: true});
  } else if (state.sheets["referral_events"] && state.sheets["referral_events"].headers.includes("timestamp") && state.sheets["referral_events"].headers.includes("eventType")) {
    // derive daily counts from referral_events
    const rows = state.sheets["referral_events"].rows;
    const m = new Map(); // date -> {share, visit, complete}
    for (const r of rows) {
      const d = safeDate(r["timestamp"]);
      if (!d) continue;
      const day = toISODateString(d);
      if (!m.has(day)) m.set(day, { share: 0, referral_visit: 0, referral_complete: 0 });
      const obj = m.get(day);
      const t = String(r["eventType"] || "");
      if (t === "share") obj.share++;
      if (t === "referral_visit") obj.referral_visit++;
      if (t === "referral_complete") obj.referral_complete++;
    }
    const days = Array.from(m.keys()).sort();
    const getSeries = (k) => days.map(d => m.get(d)[k] || 0);
    const traces = [
      {type:"scatter", mode:"lines+markers", name:"share", x: days, y: getSeries("share")},
      {type:"scatter", mode:"lines+markers", name:"referral_visit", x: days, y: getSeries("referral_visit")},
      {type:"scatter", mode:"lines+markers", name:"referral_complete", x: days, y: getSeries("referral_complete")},
    ];
    Plotly.newPlot("plotEventsDaily", traces, {
      margin: {l: 40, r: 10, t: 10, b: 40},
      legend: {orientation: "h"}
    }, {displayModeBar: false, responsive: true});
  } else {
    $("plotEventsDaily").innerHTML = `<div class="text-muted small">events_daily か referral_events(timestamp/eventType) が見つかりません</div>`;
  }

  // Sankey
  const visitsSheet = state.sheets["sankey_visits"] ? "sankey_visits" : (state.sheets["sankey_visits_acyclic"] ? "sankey_visits_acyclic" : null);
  const completesSheet = state.sheets["sankey_completes"] ? "sankey_completes" : (state.sheets["sankey_completes_acyclic"] ? "sankey_completes_acyclic" : null);

  if (visitsSheet) drawSankey("plotSankeyVisits", visitsSheet);
  else $("plotSankeyVisits").innerHTML = `<div class="text-muted small">sankey_visits が見つかりません</div>`;

  if (completesSheet) drawSankey("plotSankeyCompletes", completesSheet);
  else $("plotSankeyCompletes").innerHTML = `<div class="text-muted small">sankey_completes が見つかりません</div>`;
}

function drawSankey(targetDivId, sheetName) {
  const rows = state.sheets[sheetName].rows;
  const headers = state.sheets[sheetName].headers;
  if (!headers.includes("source") || !headers.includes("target") || !headers.includes("value")) {
    $(targetDivId).innerHTML = `<div class="text-muted small">必要な列（source/target/value）がありません</div>`;
    return;
  }
  const nodes = new Map(); // label->index
  const addNode = (label) => {
    const key = String(label);
    if (!nodes.has(key)) nodes.set(key, nodes.size);
    return nodes.get(key);
  };
  const src = [];
  const tgt = [];
  const val = [];
  for (const r of rows) {
    if (isMissing(r.source) || isMissing(r.target)) continue;
    const s = addNode(r.source);
    const t = addNode(r.target);
    const v = safeNumber(r.value) || 0;
    src.push(s); tgt.push(t); val.push(v);
  }
  const labels = Array.from(nodes.keys());
  const data = [{
    type: "sankey",
    orientation: "h",
    node: { label: labels, pad: 15, thickness: 15 },
    link: { source: src, target: tgt, value: val }
  }];
  Plotly.newPlot(targetDivId, data, {
    margin: {l: 10, r: 10, t: 10, b: 10},
  }, {displayModeBar: false, responsive: true});
}

/**
 * ============================
 * Referral Deep Dive
 * ============================
 */

function getReferralEventsAll() {
  const sheet = state.sheets["referral_events"];
  if (!sheet) return [];
  const rows = sheet.rows || [];
  const hasPayload = (sheet.headers || []).includes("payload_json");
  const events = [];
  for (const r of rows) {
    const ts = safeDate(r["timestamp"] || r["createdAt"] || r["date"]);
    if (!ts) continue;
    const eventType = String(r["eventType"] || "").trim();
    const payload = hasPayload ? parseJsonSafe(r["payload_json"]) : {};
    const userId = r["userId"] || payload.userId || null;
    const referrerId = r["referrerId"] || payload.referrerId || null;

    events.push({
      ts,
      eventType,
      userId,
      referrerId,
      platform: payload.platform || r["platform"] || null,
      userType: payload.userType || r["userType"] || null,
      userName: payload.userName || r["userName"] || null,
      userEmail: payload.userEmail || r["userEmail"] || null,
      gender: payload.gender || r["gender"] || null,
    });
  }
  // sort by time
  events.sort((a, b) => a.ts - b.ts);
  return events;
}

function computeEventDateBounds(events) {
  if (!events || events.length === 0) return { min: null, max: null };
  return { min: events[0].ts, max: events[events.length - 1].ts };
}

function filterEventsByRange(events, startDate, endDate) {
  // startDate/endDate: Date (local midnight). end is inclusive.
  const endExcl = endDate ? addDays(endDate, 1) : null;
  return (events || []).filter(e => {
    if (startDate && e.ts < startDate) return false;
    if (endExcl && e.ts >= endExcl) return false;
    return true;
  });
}

function buildEdgesFromEvents(eventsFiltered) {
  // returns {edges:[], edgesByReferrer: Map}
  const edgesMap = new Map(); // key = referrerId||userId
  const keyOf = (rid, uid) => `${rid}||${uid}`;
  for (const e of eventsFiltered) {
    if (e.eventType !== "referral_visit" && e.eventType !== "referral_complete") continue;
    if (!e.referrerId || !e.userId) continue;
    const k = keyOf(e.referrerId, e.userId);
    if (!edgesMap.has(k)) {
      edgesMap.set(k, {
        referrerId: e.referrerId,
        userId: e.userId,
        visits: 0,
        completes: 0,
        first_visit_at: null,
        first_complete_at: null,
        hours_to_complete: null,
      });
    }
    const obj = edgesMap.get(k);
    if (e.eventType === "referral_visit") {
      obj.visits += 1;
      if (!obj.first_visit_at || e.ts < obj.first_visit_at) obj.first_visit_at = e.ts;
    }
    if (e.eventType === "referral_complete") {
      obj.completes += 1;
      if (!obj.first_complete_at || e.ts < obj.first_complete_at) obj.first_complete_at = e.ts;
    }
  }

  const edges = Array.from(edgesMap.values()).map(o => {
    let hours = null;
    if (o.first_visit_at && o.first_complete_at) {
      hours = (o.first_complete_at.getTime() - o.first_visit_at.getTime()) / 3600000;
      if (!Number.isFinite(hours)) hours = null;
    }
    return {
      referrerId: o.referrerId,
      userId: o.userId,
      visits: o.visits,
      completes: o.completes,
      first_visit_at: o.first_visit_at ? o.first_visit_at.toISOString() : null,
      first_complete_at: o.first_complete_at ? o.first_complete_at.toISOString() : null,
      hours_to_complete: hours,
    };
  });

  const byRef = new Map();
  for (const e of edges) {
    if (!byRef.has(e.referrerId)) byRef.set(e.referrerId, []);
    byRef.get(e.referrerId).push(e);
  }
  for (const [rid, arr] of byRef.entries()) {
    arr.sort((a, b) => (safeDate(a.first_visit_at) || 0) - (safeDate(b.first_visit_at) || 0));
  }
  return { edges, edgesByReferrer: byRef };
}

function buildUserInfoMap(eventsAll) {
  const m = new Map(); // userId -> info
  const upsert = (userId, patch) => {
    if (!userId) return;
    if (!m.has(userId)) m.set(userId, {});
    const obj = m.get(userId);
    for (const [k, v] of Object.entries(patch)) {
      if (isMissing(v)) continue;
      if (obj[k] === undefined || obj[k] === null || obj[k] === "") obj[k] = v;
    }
  };

  for (const e of eventsAll) {
    if (e.eventType === "share" && e.userId) {
      upsert(e.userId, { userType: e.userType, gender: e.gender, userName: e.userName, userEmail: e.userEmail });
    }
    if (e.eventType === "referral_complete" && e.userId) {
      // completing user's info (not necessarily needed, but can be useful)
      upsert(e.userId, { userType: e.userType, gender: e.gender, userName: e.userName, userEmail: e.userEmail });
    }
  }
  return m;
}

function computeReferrerStats(eventsAll, eventsFiltered) {
  const userInfo = buildUserInfoMap(eventsAll);

  const stats = new Map(); // referrerId -> obj
  const get = (rid) => {
    if (!stats.has(rid)) {
      const info = userInfo.get(rid) || {};
      stats.set(rid, {
        referrerId: rid,
        referrerType: info.userType || null,
        referrerGender: info.gender || null,

        shares_total: 0,
        shares_line: 0,
        shares_twitter: 0,
        shares_copy: 0,

        visitors_set: new Set(),
        completes_set: new Set(),

        unique_invited_visitors: 0,
        unique_invited_completes: 0,
        visit_to_complete_rate: null,
        avg_hours_to_complete: null,
        median_hours_to_complete: null,
      });
    }
    return stats.get(rid);
  };

  // from referrer_summary (optional): include referrers that might have 0 events in filtered window
  if (state.sheets["referrer_summary"]) {
    const rs = state.sheets["referrer_summary"].rows || [];
    for (const r of rs) {
      const rid = r["referrerId"];
      if (isMissing(rid)) continue;
      get(String(rid));
    }
  }

  // aggregate filtered events
  for (const e of eventsFiltered) {
    if (e.eventType === "share" && e.userId) {
      const s = get(String(e.userId));
      s.shares_total += 1;
      const p = String(e.platform || "").toLowerCase();
      if (p === "line") s.shares_line += 1;
      else if (p === "twitter") s.shares_twitter += 1;
      else if (p === "copy") s.shares_copy += 1;
    }
    if (e.eventType === "referral_visit" && e.referrerId && e.userId) {
      const s = get(String(e.referrerId));
      s.visitors_set.add(String(e.userId));
    }
    if (e.eventType === "referral_complete" && e.referrerId && e.userId) {
      const s = get(String(e.referrerId));
      s.completes_set.add(String(e.userId));
    }
  }

  // edges + time-to-complete from filtered events
  const { edges, edgesByReferrer } = buildEdgesFromEvents(eventsFiltered);

  for (const [rid, s] of stats.entries()) {
    s.unique_invited_visitors = s.visitors_set.size;
    s.unique_invited_completes = s.completes_set.size;
    s.visit_to_complete_rate = (s.unique_invited_visitors > 0) ? (s.unique_invited_completes / s.unique_invited_visitors) : null;

    const hrs = (edgesByReferrer.get(rid) || [])
      .map(e => e.hours_to_complete)
      .filter(v => v !== null && v !== undefined && Number.isFinite(v))
      .sort((a, b) => a - b);
    if (hrs.length) {
      s.avg_hours_to_complete = hrs.reduce((a, b) => a + b, 0) / hrs.length;
      s.median_hours_to_complete = median(hrs);
    } else {
      s.avg_hours_to_complete = null;
      s.median_hours_to_complete = null;
    }

    // cleanup sets for display
    delete s.visitors_set;
    delete s.completes_set;
  }

  const list = Array.from(stats.values()).sort((a, b) => {
    const ca = safeNumber(a.unique_invited_completes) || 0;
    const cb = safeNumber(b.unique_invited_completes) || 0;
    if (cb !== ca) return cb - ca;
    const va = safeNumber(a.unique_invited_visitors) || 0;
    const vb = safeNumber(b.unique_invited_visitors) || 0;
    if (vb !== va) return vb - va;
    const sa = safeNumber(a.shares_total) || 0;
    const sb = safeNumber(b.shares_total) || 0;
    return sb - sa;
  });

  return { list, edges, edgesByReferrer, userInfo };
}

function computeDailyEventCounts(eventsFiltered) {
  const m = new Map(); // day -> {share, visit, complete}
  for (const e of eventsFiltered) {
    const day = toISODateString(e.ts);
    if (!m.has(day)) m.set(day, { share: 0, referral_visit: 0, referral_complete: 0 });
    const obj = m.get(day);
    if (e.eventType === "share") obj.share += 1;
    if (e.eventType === "referral_visit") obj.referral_visit += 1;
    if (e.eventType === "referral_complete") obj.referral_complete += 1;
  }
  const days = Array.from(m.keys()).sort();
  return { days, m };
}

function computeSharePlatformCounts(eventsFiltered) {
  const m = new Map(); // platform -> count
  for (const e of eventsFiltered) {
    if (e.eventType !== "share") continue;
    const p = String(e.platform || "unknown");
    m.set(p, (m.get(p) || 0) + 1);
  }
  const arr = Array.from(m.entries()).sort((a, b) => b[1] - a[1]);
  return arr;
}

function computeTypePerformance(referrerStatsList) {
  // returns [{type, referrers, shares, visitors, completes, v2c_rate, avg_hours}]
  const m = new Map();
  for (const r of referrerStatsList) {
    const t = r.referrerType ? String(r.referrerType) : "unknown";
    if (!m.has(t)) m.set(t, { type: t, referrers: 0, shares: 0, visitors: 0, completes: 0, hours: [] });
    const o = m.get(t);
    o.referrers += 1;
    o.shares += safeNumber(r.shares_total) || 0;
    o.visitors += safeNumber(r.unique_invited_visitors) || 0;
    o.completes += safeNumber(r.unique_invited_completes) || 0;
    if (r.avg_hours_to_complete !== null && r.avg_hours_to_complete !== undefined && Number.isFinite(r.avg_hours_to_complete)) {
      o.hours.push(r.avg_hours_to_complete);
    }
  }
  const out = Array.from(m.values()).map(o => {
    const v2c = o.visitors > 0 ? o.completes / o.visitors : null;
    const avgH = o.hours.length ? (o.hours.reduce((a, b) => a + b, 0) / o.hours.length) : null;
    return { type: o.type, referrers: o.referrers, shares: o.shares, visitors: o.visitors, completes: o.completes, v2c_rate: v2c, avg_hours: avgH };
  }).sort((a, b) => (safeNumber(b.completes) || 0) - (safeNumber(a.completes) || 0));
  return out;
}

function computeReferralDiagnosisJoin(eventsFiltered, startDate, endDate) {
  // Join referral_complete(payload.userEmail) with diagnosis.email
  const diagSheet = state.sheets["diagnosis"];
  if (!diagSheet) return null;
  const headers = diagSheet.headers || [];
  if (!headers.includes("email")) return null;

  const hasCreatedAt = headers.includes("createdAt");
  const hasType = headers.includes("type");
  const hasInterested = headers.includes("interested");

  const endExcl = endDate ? addDays(endDate, 1) : null;

  // Build latest diagnosis row per email (within period if createdAt exists)
  const latestByEmail = new Map(); // email -> {createdAt, type, interested}
  for (const r of (diagSheet.rows || [])) {
    const email = r["email"];
    if (isMissing(email)) continue;
    const em = String(email).trim();
    if (!em) continue;

    let d = hasCreatedAt ? safeDate(r["createdAt"]) : new Date(0);
    if (hasCreatedAt) {
      if (!d) continue;
      if (startDate && d < startDate) continue;
      if (endExcl && d >= endExcl) continue;
    }

    const prev = latestByEmail.get(em);
    if (!prev || (d && prev.createdAt && d > prev.createdAt)) {
      latestByEmail.set(em, {
        email: em,
        createdAt: d,
        type: hasType ? (isMissing(r["type"]) ? null : String(r["type"])) : null,
        interested: hasInterested ? r["interested"] : null,
      });
    }
  }

  // Referred (unique emails) from referral_complete in the filtered period
  const referredEmails = new Set();
  for (const e of (eventsFiltered || [])) {
    if (e.eventType !== "referral_complete") continue;
    if (isMissing(e.userEmail)) continue;
    const em = String(e.userEmail).trim();
    if (!em) continue;
    referredEmails.add(em);
  }

  const matched = [];
  for (const em of referredEmails) {
    const row = latestByEmail.get(em);
    if (row) matched.push(row);
  }

  const overallRows = Array.from(latestByEmail.values());

  const countTypes = (rows) => {
    const m = new Map();
    for (const r of rows) {
      const t = r.type || "unknown";
      m.set(t, (m.get(t) || 0) + 1);
    }
    return m;
  };

  const isInterested = (v) => {
    return String(v) === "1" || v === 1 || v === true;
  };

  let refInterested = 0;
  for (const r of matched) if (isInterested(r.interested)) refInterested++;

  let allInterested = 0;
  for (const r of overallRows) if (isInterested(r.interested)) allInterested++;

  const refRate = matched.length ? (refInterested / matched.length) : null;
  const allRate = overallRows.length ? (allInterested / overallRows.length) : null;

  return {
    referred_unique_emails: referredEmails.size,
    matched_unique_emails: matched.length,
    refRate,
    allRate,
    refType: countTypes(matched),
    allType: countTypes(overallRows),
  };
}

function renderReferralDiagnosisJoin(join) {
  const el1 = $("kpiReferredUniqueEmail");
  const el2 = $("kpiReferredMatched");
  const el3 = $("kpiReferredInterested");
  const plotId = "plotReferredTypeCompare";

  if (!el1 || !el2 || !el3 || !$(plotId)) return;

  if (!join) {
    el1.textContent = "-";
    el2.textContent = "-";
    el3.textContent = "-";
    $(plotId).innerHTML = `<div class="text-muted small">diagnosis（email/type/createdAt）が無いので突合できません</div>`;
    return;
  }

  el1.textContent = formatInt(join.referred_unique_emails);
  el2.textContent = formatInt(join.matched_unique_emails);

  const a = (join.refRate === null) ? "-" : formatPct(join.refRate, 1);
  const b = (join.allRate === null) ? "-" : formatPct(join.allRate, 1);
  el3.textContent = `${a} / ${b}`;

  // Build union of types: top overall + all referred
  const allCounts = {};
  for (const [k, v] of join.allType.entries()) allCounts[k] = v;
  const refCounts = {};
  for (const [k, v] of join.refType.entries()) refCounts[k] = v;

  const topAll = Object.entries(allCounts).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([t]) => t);
  const typesSet = new Set(topAll);
  Object.keys(refCounts).forEach(t => typesSet.add(t));

  const x = Array.from(typesSet);
  x.sort((a, b) => (allCounts[b] || 0) - (allCounts[a] || 0));

  Plotly.newPlot(plotId, [
    { type: "bar", name: "referred_complete", x, y: x.map(t => refCounts[t] || 0) },
    { type: "bar", name: "all_diagnosis", x, y: x.map(t => allCounts[t] || 0) },
  ], {
    margin: { l: 40, r: 10, t: 10, b: 120 },
    barmode: "group",
    xaxis: { tickangle: -45 },
    yaxis: { title: "users (unique email)" },
    legend: { orientation: "h" },
  }, { displayModeBar: false, responsive: true });
}

function renderReferralDeepDive() {
  const empty = $("referralEmpty");
  const content = $("referralContent");
  if (!empty || !content) return;

  if (!state.workbook) {
    empty.classList.remove("d-none");
    content.classList.add("d-none");
    return;
  }

  if (!state.sheets["referral_events"]) {
    empty.classList.remove("d-none");
    empty.innerHTML = `紹介系シート（referral_events）が見つかりません。<br/>「シート探索」タブで利用可能なシートを確認してください。`;
    content.classList.add("d-none");
    return;
  }

  // Base events
  const eventsAll = getReferralEventsAll();
  const bounds = computeEventDateBounds(eventsAll);

  empty.classList.add("d-none");
  content.classList.remove("d-none");

  // Init date inputs (only if empty)
  const startInput = $("refStartDate");
  const endInput = $("refEndDate");
  if (bounds.min && bounds.max) {
    const minStr = toISODateString(bounds.min);
    const maxStr = toISODateString(bounds.max);
    if (startInput && !startInput.value) startInput.value = minStr;
    if (endInput && !endInput.value) endInput.value = maxStr;
  }

  // filter
  let start = startInput ? parseDateInput(startInput.value) : null;
  let end = endInput ? parseDateInput(endInput.value) : null;
  if (start && end && start > end) {
    // swap
    const tmp = start; start = end; end = tmp;
    if (startInput) startInput.value = toISODateString(start);
    if (endInput) endInput.value = toISODateString(end);
  }

  const eventsFiltered = filterEventsByRange(eventsAll, start, end);

  // compute
  const { list: referrerStats, edges, edgesByReferrer } = computeReferrerStats(eventsAll, eventsFiltered);
  const daily = computeDailyEventCounts(eventsFiltered);
  const platformCounts = computeSharePlatformCounts(eventsFiltered);
  const typePerf = computeTypePerformance(referrerStats);

  state.referralDerived = {
    bounds,
    start,
    end,
    eventsAll,
    eventsFiltered,
    referrerStats,
    edges,
    edgesByReferrer,
    daily,
    platformCounts,
    typePerf,
  };

  // Hint
  const hint = $("refFilterHint");
  if (hint) {
    const a = start ? toISODateString(start) : "-";
    const b = end ? toISODateString(end) : "-";
    const total = eventsAll.length;
    const filtered = eventsFiltered.length;
    hint.textContent = `対象期間: ${a} 〜 ${b} / events: ${filtered.toLocaleString()}（全${total.toLocaleString()}）`;
  }

  // KPI
  const shares = eventsFiltered.filter(e => e.eventType === "share").length;
  const visits = eventsFiltered.filter(e => e.eventType === "referral_visit").length;
  const completes = eventsFiltered.filter(e => e.eventType === "referral_complete").length;
  $("kpiRefShare").textContent = formatInt(shares);
  $("kpiRefVisit").textContent = formatInt(visits);
  $("kpiRefComplete").textContent = formatInt(completes);
  $("kpiRefRateSV").textContent = shares > 0 ? formatPct(visits / shares, 1) : "-";
  $("kpiRefRateVC").textContent = visits > 0 ? formatPct(completes / visits, 1) : "-";
  $("kpiRefRateSC").textContent = shares > 0 ? formatPct(completes / shares, 1) : "-";

  // plots
  renderReferralPlots();

  // referral_complete × diagnosis join
  const join = computeReferralDiagnosisJoin(eventsFiltered, start, end);
  state.referralDerived.join = join;
  renderReferralDiagnosisJoin(join);

  // leaderboard table
  renderReferrerTable();

  // populate referrer select + details
  populateReferrerSelect();
  renderReferrerDetail();

  // sankey
  drawReferralSankey();
}

function renderReferralPlots() {
  const d = state.referralDerived;
  if (!d) return;

  // time series
  const days = d.daily.days;
  const getSeries = (k) => days.map(day => d.daily.m.get(day)[k] || 0);
  if (days.length) {
    Plotly.newPlot("plotRefTimeSeries", [
      { type: "scatter", mode: "lines+markers", name: "share", x: days, y: getSeries("share") },
      { type: "scatter", mode: "lines+markers", name: "referral_visit", x: days, y: getSeries("referral_visit") },
      { type: "scatter", mode: "lines+markers", name: "referral_complete", x: days, y: getSeries("referral_complete") },
    ], {
      margin: { l: 40, r: 10, t: 10, b: 40 },
      legend: { orientation: "h" }
    }, { displayModeBar: false, responsive: true });
  } else {
    $("plotRefTimeSeries").innerHTML = `<div class="text-muted small">対象期間にイベントがありません</div>`;
  }

  // platform
  if (d.platformCounts.length) {
    const x = d.platformCounts.map(([p]) => p);
    const y = d.platformCounts.map(([, c]) => c);
    Plotly.newPlot("plotSharePlatform", [{
      type: "bar", x, y
    }], {
      margin: { l: 40, r: 10, t: 10, b: 80 },
      xaxis: { tickangle: -30 }
    }, { displayModeBar: false, responsive: true });
  } else {
    $("plotSharePlatform").innerHTML = `<div class="text-muted small">shareイベントがありません</div>`;
  }

  // type performance (v2c rate)
  if (d.typePerf.length) {
    const top = d.typePerf.slice(0, 20);
    const x = top.map(o => o.type);
    const y = top.map(o => o.v2c_rate === null ? 0 : o.v2c_rate);
    const text = top.map(o => `referrers=${o.referrers}, visitors=${o.visitors}, completes=${o.completes}`);
    Plotly.newPlot("plotTypePerformance", [{
      type: "bar",
      x, y,
      text,
      hovertemplate: "%{x}<br>visit→complete=%{y:.2%}<br>%{text}<extra></extra>"
    }], {
      margin: { l: 50, r: 10, t: 10, b: 120 },
      yaxis: { tickformat: ".0%" },
      xaxis: { tickangle: -45 }
    }, { displayModeBar: false, responsive: true });
  } else {
    $("plotTypePerformance").innerHTML = `<div class="text-muted small">type情報が不足しています</div>`;
  }

  // scatter visitors vs completes
  const flags = getUIFlags();
  if (d.referrerStats.length) {
    const xs = d.referrerStats.map(r => safeNumber(r.unique_invited_visitors) || 0);
    const ys = d.referrerStats.map(r => safeNumber(r.unique_invited_completes) || 0);
    const sizes = d.referrerStats.map(r => Math.max(6, Math.sqrt(safeNumber(r.shares_total) || 0) * 6));
    const labels = d.referrerStats.map(r => {
      const id = flags.maskPII ? maskIdLike(r.referrerId) : r.referrerId;
      const t = r.referrerType ? ` / ${r.referrerType}` : "";
      return id + t;
    });
    Plotly.newPlot("plotRefScatter", [{
      type: "scatter",
      mode: "markers",
      x: xs,
      y: ys,
      text: labels,
      marker: { size: sizes, sizemode: "area", opacity: 0.75 },
      hovertemplate: "%{text}<br>visitors=%{x}<br>completes=%{y}<extra></extra>"
    }], {
      margin: { l: 50, r: 10, t: 10, b: 40 },
      xaxis: { title: "unique invited visitors" },
      yaxis: { title: "unique invited completes" }
    }, { displayModeBar: false, responsive: true });
  } else {
    $("plotRefScatter").innerHTML = `<div class="text-muted small">referrerがいません</div>`;
  }
}

function renderReferrerTable() {
  const d = state.referralDerived;
  if (!d) return;
  const flags = getUIFlags();

  const table = $("referrerTable");
  if (!table) return;

  const cols = [
    "referrerId",
    "referrerType",
    "shares_total",
    "shares_line",
    "shares_copy",
    "shares_twitter",
    "unique_invited_visitors",
    "unique_invited_completes",
    "visit_to_complete_rate",
    "avg_hours_to_complete",
  ];

  // build thead/tbody
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  for (const c of cols) {
    const th = document.createElement("th");
    th.textContent = c;
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  for (const r of d.referrerStats) {
    const tr = document.createElement("tr");
    tr.setAttribute("data-referrer-id", r.referrerId);

    for (const c of cols) {
      const td = document.createElement("td");
      let v = r[c];

      if (c === "referrerId") v = flags.maskPII ? maskIdLike(v) : v;
      if (c === "visit_to_complete_rate") v = (v === null ? "" : formatPct(v, 1));
      if (c === "avg_hours_to_complete") v = (v === null ? "" : formatHours(v, 2));
      if (["shares_total","shares_line","shares_copy","shares_twitter","unique_invited_visitors","unique_invited_completes"].includes(c)) v = formatInt(v);

      td.textContent = isMissing(v) ? "" : String(v);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  if (state.referrerTable) {
    try { state.referrerTable.destroy(); } catch(e) {}
    state.referrerTable = null;
  }
  state.referrerTable = new DataTable(table, {
    paging: true,
    pageLength: 25,
    searching: true,
    info: true,
    responsive: true,
    order: [[7, "desc"]],
  });
}

function populateReferrerSelect() {
  const d = state.referralDerived;
  if (!d) return;
  const flags = getUIFlags();
  const sel = $("referrerSelect");
  if (!sel) return;

  const prev = state.referralSelectedReferrer;

  sel.innerHTML = "";
  for (const r of d.referrerStats) {
    const opt = document.createElement("option");
    opt.value = r.referrerId;
    const id = flags.maskPII ? maskIdLike(r.referrerId) : r.referrerId;
    const t = r.referrerType ? ` / ${r.referrerType}` : "";
    const c = safeNumber(r.unique_invited_completes) || 0;
    opt.textContent = `${id}${t} (complete=${c})`;
    sel.appendChild(opt);
  }

  // choose selection
  if (prev && d.referrerStats.some(r => r.referrerId === prev)) {
    sel.value = prev;
    state.referralSelectedReferrer = prev;
  } else {
    const first = d.referrerStats[0]?.referrerId || null;
    if (first) {
      sel.value = first;
      state.referralSelectedReferrer = first;
    }
  }
}

function renderReferrerDetail() {
  const d = state.referralDerived;
  if (!d) return;

  const sel = $("referrerSelect");
  if (!sel) return;

  const rid = state.referralSelectedReferrer || sel.value;
  if (!rid) return;
  state.referralSelectedReferrer = rid;
  sel.value = rid;

  const flags = getUIFlags();

  const stats = d.referrerStats.find(r => r.referrerId === rid);
  const edges = d.edgesByReferrer.get(rid) || [];
  const events = d.eventsFiltered.filter(e => e.userId === rid || e.referrerId === rid);

  // KPI cards
  const kpi = $("referrerDetailKpis");
  if (kpi) {
    const idDisp = flags.maskPII ? maskIdLike(rid) : rid;
    const typeDisp = stats?.referrerType ? String(stats.referrerType) : "-";
    const share = stats ? stats.shares_total : 0;
    const vis = stats ? stats.unique_invited_visitors : 0;
    const comp = stats ? stats.unique_invited_completes : 0;
    const v2c = stats ? stats.visit_to_complete_rate : null;
    const avgH = stats ? stats.avg_hours_to_complete : null;
    const medH = stats ? stats.median_hours_to_complete : null;

    kpi.innerHTML = `
      <div class="col-md-3"><div class="card h-100"><div class="card-body">
        <div class="small text-muted">referrer</div><div class="h5 mb-0">${idDisp}</div>
        <div class="small text-muted mt-1">type: ${typeDisp}</div>
      </div></div></div>
      <div class="col-md-3"><div class="card h-100"><div class="card-body">
        <div class="small text-muted">shares</div><div class="h5 mb-0">${formatInt(share)}</div>
      </div></div></div>
      <div class="col-md-3"><div class="card h-100"><div class="card-body">
        <div class="small text-muted">visitors / completes</div><div class="h5 mb-0">${formatInt(vis)} / ${formatInt(comp)}</div>
        <div class="small text-muted mt-1">visit→complete: ${v2c === null ? "-" : formatPct(v2c, 1)}</div>
      </div></div></div>
      <div class="col-md-3"><div class="card h-100"><div class="card-body">
        <div class="small text-muted">hours to complete</div><div class="h5 mb-0">${avgH === null ? "-" : formatHours(avgH, 2)}</div>
        <div class="small text-muted mt-1">median: ${medH === null ? "-" : formatHours(medH, 2)}</div>
      </div></div></div>
    `;
  }

  // timeline: daily counts for this referrer (share/visit/complete)
  const m = new Map();
  for (const e of events) {
    const day = toISODateString(e.ts);
    if (!m.has(day)) m.set(day, { share: 0, referral_visit: 0, referral_complete: 0 });
    const obj = m.get(day);
    if (e.eventType === "share" && e.userId === rid) obj.share += 1;
    if (e.eventType === "referral_visit" && e.referrerId === rid) obj.referral_visit += 1;
    if (e.eventType === "referral_complete" && e.referrerId === rid) obj.referral_complete += 1;
  }
  const days = Array.from(m.keys()).sort();
  if (days.length) {
    Plotly.newPlot("plotReferrerTimeline", [
      { type: "scatter", mode: "lines+markers", name: "share", x: days, y: days.map(dy => m.get(dy).share) },
      { type: "scatter", mode: "lines+markers", name: "referral_visit", x: days, y: days.map(dy => m.get(dy).referral_visit) },
      { type: "scatter", mode: "lines+markers", name: "referral_complete", x: days, y: days.map(dy => m.get(dy).referral_complete) },
    ], {
      margin: { l: 40, r: 10, t: 10, b: 40 },
      legend: { orientation: "h" }
    }, { displayModeBar: false, responsive: true });
  } else {
    $("plotReferrerTimeline").innerHTML = `<div class="text-muted small">対象期間にイベントがありません</div>`;
  }

  // hours histogram
  const hrs = edges.map(e => e.hours_to_complete).filter(v => v !== null && v !== undefined && Number.isFinite(v));
  if (hrs.length) {
    Plotly.newPlot("plotHoursToComplete", [{
      type: "histogram",
      x: hrs,
      nbinsx: Math.min(40, Math.max(10, Math.round(Math.sqrt(hrs.length))))
    }], {
      margin: { l: 50, r: 10, t: 10, b: 40 },
      xaxis: { title: "hours_to_complete" },
      yaxis: { title: "count" }
    }, { displayModeBar: false, responsive: true });
  } else {
    $("plotHoursToComplete").innerHTML = `<div class="text-muted small">completeがありません</div>`;
  }

  // edges table
  renderRefEdgesTable(edges);
}

function renderRefEdgesTable(edges) {
  const flags = getUIFlags();
  const table = $("refEdgesTable");
  if (!table) return;

  const cols = ["userId", "visits", "completes", "first_visit_at", "first_complete_at", "hours_to_complete"];

  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  for (const c of cols) {
    const th = document.createElement("th");
    th.textContent = c;
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  for (const e of edges) {
    const tr = document.createElement("tr");
    for (const c of cols) {
      const td = document.createElement("td");
      let v = e[c];

      if (c === "userId") v = flags.maskPII ? maskIdLike(v) : v;
      if (c === "first_visit_at" || c === "first_complete_at") {
        const d = safeDate(v);
        v = d ? d.toISOString().replace("T", " ").slice(0, 19) : "";
      }
      if (c === "hours_to_complete") v = (v === null ? "" : formatHours(v, 2));
      if (c === "visits" || c === "completes") v = formatInt(v);

      td.textContent = isMissing(v) ? "" : String(v);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  if (state.refEdgesTable) {
    try { state.refEdgesTable.destroy(); } catch(e) {}
    state.refEdgesTable = null;
  }
  state.refEdgesTable = new DataTable(table, {
    paging: true,
    pageLength: 25,
    searching: true,
    info: true,
    responsive: true,
    order: [[5, "asc"]],
  });
}

function pickSankeySheet(kind, variant) {
  // returns sheetName or null
  if (kind === "visits") {
    const ac = state.sheets["sankey_visits_acyclic"] ? "sankey_visits_acyclic" : null;
    const raw = state.sheets["sankey_visits"] ? "sankey_visits" : null;
    if (variant === "raw") return raw || ac;
    return ac || raw;
  }
  if (kind === "completes") {
    const ac = state.sheets["sankey_completes_acyclic"] ? "sankey_completes_acyclic" : null;
    const raw = state.sheets["sankey_completes"] ? "sankey_completes" : null;
    if (variant === "raw") return raw || ac;
    return ac || raw;
  }
  return null;
}

function drawReferralSankey() {
  const kindSel = $("refSankeyKind");
  const varSel = $("refSankeyVariant");
  const minEl = $("refSankeyMinValue");
  if (!kindSel || !varSel || !minEl) return;

  const kind = kindSel.value;
  const variant = varSel.value;
  const sheetName = pickSankeySheet(kind, variant);

  if (!sheetName) {
    $("plotReferralSankey").innerHTML = `<div class="text-muted small">Sankey用のシート（${kind}）が見つかりません</div>`;
    const st = $("refNetworkStats");
    if (st) st.textContent = "-";
    return;
  }

  const minValue = safeNumber(minEl.value) || 0;
  drawSankeyFiltered("plotReferralSankey", sheetName, minValue);

  // stats
  const stats = computeNetworkStatsFromSankeySheet(sheetName, minValue);
  const st = $("refNetworkStats");
  if (st) {
    if (!stats) {
      st.textContent = "-";
    } else {
      const depth = (stats.longestPathLength === null) ? "n/a" : String(stats.longestPathLength);
      st.textContent = `sheet=${sheetName} / nodes=${stats.nodes} / edges=${stats.edges} / depth(max edges)=${depth} / max outdegree=${stats.maxOutdegree}`;
    }
  }
}

function drawSankeyFiltered(targetDivId, sheetName, minValue = 0) {
  const rows = state.sheets[sheetName].rows;
  const headers = state.sheets[sheetName].headers;
  if (!headers.includes("source") || !headers.includes("target") || !headers.includes("value")) {
    $(targetDivId).innerHTML = `<div class="text-muted small">必要な列（source/target/value）がありません</div>`;
    return;
  }
  const nodes = new Map(); // label->index
  const addNode = (label) => {
    const key = String(label);
    if (!nodes.has(key)) nodes.set(key, nodes.size);
    return nodes.get(key);
  };
  const src = [];
  const tgt = [];
  const val = [];
  for (const r of rows) {
    if (isMissing(r.source) || isMissing(r.target)) continue;
    const v = safeNumber(r.value) || 0;
    if (v < minValue) continue;
    const s = addNode(r.source);
    const t = addNode(r.target);
    src.push(s); tgt.push(t); val.push(v);
  }
  const labels = Array.from(nodes.keys());
  const data = [{
    type: "sankey",
    orientation: "h",
    node: { label: labels, pad: 15, thickness: 15 },
    link: { source: src, target: tgt, value: val }
  }];
  Plotly.newPlot(targetDivId, data, {
    margin: { l: 10, r: 10, t: 10, b: 10 },
  }, { displayModeBar: false, responsive: true });
}

function computeNetworkStatsFromSankeySheet(sheetName, minValue = 0) {
  const sheet = state.sheets[sheetName];
  if (!sheet) return null;
  const rows = sheet.rows || [];
  const edges = [];
  const nodesSet = new Set();
  const outdeg = new Map();

  for (const r of rows) {
    if (isMissing(r.source) || isMissing(r.target)) continue;
    const v = safeNumber(r.value) || 0;
    if (v < minValue) continue;
    const s = String(r.source);
    const t = String(r.target);
    edges.push([s, t]);
    nodesSet.add(s); nodesSet.add(t);
    outdeg.set(s, (outdeg.get(s) || 0) + 1);
  }

  const nodes = Array.from(nodesSet);
  const maxOutdegree = nodes.reduce((mx, n) => Math.max(mx, outdeg.get(n) || 0), 0);

  // longest path (only if DAG)
  const longest = dagLongestPathLength(nodes, edges);

  return { nodes: nodes.length, edges: edges.length, maxOutdegree, longestPathLength: longest };
}

function dagLongestPathLength(nodes, edges) {
  // edges: array of [s,t]. returns length in edges, or null if cycle.
  const indeg = new Map();
  const adj = new Map();
  for (const n of nodes) { indeg.set(n, 0); adj.set(n, []); }
  for (const [s, t] of edges) {
    if (!adj.has(s)) { adj.set(s, []); indeg.set(s, indeg.get(s) || 0); }
    if (!adj.has(t)) { adj.set(t, []); indeg.set(t, indeg.get(t) || 0); }
    adj.get(s).push(t);
    indeg.set(t, (indeg.get(t) || 0) + 1);
  }

  const q = [];
  for (const [n, d] of indeg.entries()) if (d === 0) q.push(n);
  const order = [];
  while (q.length) {
    const n = q.shift();
    order.push(n);
    for (const t of adj.get(n) || []) {
      indeg.set(t, indeg.get(t) - 1);
      if (indeg.get(t) === 0) q.push(t);
    }
  }
  if (order.length !== indeg.size) return null; // cycle

  const dp = new Map();
  for (const n of order) dp.set(n, 0);
  let best = 0;
  for (const n of order) {
    const base = dp.get(n) || 0;
    for (const t of adj.get(n) || []) {
      const cand = base + 1;
      if ((dp.get(t) || 0) < cand) dp.set(t, cand);
      if (cand > best) best = cand;
    }
  }
  return best;
}

function setCurrentSheet(sheetName) {
  state.currentSheet = sheetName;
  // update selects
  $("sheetSelect").value = sheetName;
  $("profileSheetSelect").value = sheetName;
  $("chartsSheetSelect").value = sheetName;
  $("exportSheetSelect").value = sheetName;
  renderSheetTable();
  renderProfile();
  syncChartColumns();
}

function populateSheetSelects() {
  const selects = ["sheetSelect", "profileSheetSelect", "chartsSheetSelect", "exportSheetSelect"].map(id => $(id));
  for (const sel of selects) {
    sel.innerHTML = "";
    for (const s of state.sheetOrder) {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      sel.appendChild(opt);
    }
  }
}

function renderSheetTable() {
  const sheetName = $("sheetSelect").value;
  const flags = getUIFlags();
  const rawRows = state.sheets[sheetName]?.rows || [];
  const headers = getVisibleHeaders(sheetName, flags);
  const displayRows = getDisplayRows(sheetName, flags);

  $("sheetMeta").textContent = `${sheetName}: rows=${rawRows.length.toLocaleString()}, cols=${headers.length}`;

  // build thead/tbody
  const table = $("dataTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  for (const h of headers) {
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  for (const r of displayRows) {
    const tr = document.createElement("tr");
    for (const h of headers) {
      const td = document.createElement("td");
      const v = r[h];
      td.textContent = isMissing(v) ? "" : String(v);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  // DataTables init/destroy
  if (state.dataTable) {
    try { state.dataTable.destroy(); } catch(e) {}
    state.dataTable = null;
  }
  state.dataTable = new DataTable(table, {
    paging: true,
    pageLength: 25,
    lengthMenu: [ [25, 50, 100, 250], [25, 50, 100, 250] ],
    searching: true,
    info: true,
    responsive: true,
    order: [],
  });
}

function renderProfile() {
  const sheetName = $("profileSheetSelect").value;
  const flags = getUIFlags();
  const rows = state.sheets[sheetName]?.rows || [];
  const headers = getVisibleHeaders(sheetName, flags);

  const tbody = $("profileTable").querySelector("tbody");
  tbody.innerHTML = "";

  for (const h of headers) {
    const colValuesRaw = rows.map(r => r[h]).filter(v => !isMissing(v)).slice(0, 2000);
    const colType = detectColumnType(colValuesRaw.slice(0, 200));
    const missing = rows.length - colValuesRaw.length;

    const unique = new Set(colValuesRaw.map(v => String(v))).size;

    let min = "", p50 = "", max = "", topStr = "";
    if (colType === "number") {
      const nums = colValuesRaw.map(v => safeNumber(v)).filter(v => v !== null).sort((a,b)=>a-b);
      if (nums.length) {
        min = nums[0].toFixed(3).replace(/\.?0+$/,"");
        p50 = median(nums).toFixed(3).replace(/\.?0+$/,"");
        max = nums[nums.length-1].toFixed(3).replace(/\.?0+$/,"");
      }
    } else if (colType === "date") {
      const ds = colValuesRaw.map(v => safeDate(v)).filter(d => d).sort((a,b)=>a-b);
      if (ds.length) {
        min = toISODateString(ds[0]);
        p50 = toISODateString(ds[Math.floor(ds.length/2)]);
        max = toISODateString(ds[ds.length-1]);
      }
    } else if (colType === "string" || colType === "boolean") {
      const top = freqTop(colValuesRaw, 5);
      topStr = top.map(t => {
        const v = t.value.length > 24 ? t.value.slice(0, 24) + "…" : t.value;
        return `${v}(${t.count})`;
      }).join(", ");
    }

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="codeLike">${h}</td>
      <td>${colType}</td>
      <td class="text-end">${missing.toLocaleString()}</td>
      <td class="text-end">${unique.toLocaleString()}</td>
      <td class="text-end">${min}</td>
      <td class="text-end">${p50}</td>
      <td class="text-end">${max}</td>
      <td>${topStr}</td>
    `;
    tbody.appendChild(tr);
  }
}

function syncChartColumns() {
  const sheetName = $("chartsSheetSelect").value;
  const flags = getUIFlags();
  const headers = getVisibleHeaders(sheetName, flags);

  const xSel = $("xCol");
  const ySel = $("yCol");
  xSel.innerHTML = "";
  ySel.innerHTML = "";

  for (const h of headers) {
    const optX = document.createElement("option");
    optX.value = h; optX.textContent = h;
    xSel.appendChild(optX);

    const optY = document.createElement("option");
    optY.value = h; optY.textContent = h;
    ySel.appendChild(optY);
  }

  autoPickColumns();
}

function autoPickColumns() {
  const sheetName = $("chartsSheetSelect").value;
  const flags = getUIFlags();
  const headers = getVisibleHeaders(sheetName, flags);
  const rows = state.sheets[sheetName]?.rows || [];
  const chartType = $("chartType").value;

  const sample = (col) => rows.map(r => r[col]).filter(v => !isMissing(v)).slice(0, 200);

  const colTypes = {};
  for (const h of headers) colTypes[h] = detectColumnType(sample(h));

  const numericCols = headers.filter(h => colTypes[h] === "number");
  const dateCols = headers.filter(h => colTypes[h] === "date");
  const stringCols = headers.filter(h => colTypes[h] === "string" || colTypes[h] === "boolean");

  const xSel = $("xCol");
  const ySel = $("yCol");

  const setIf = (sel, col) => {
    if (!col) return;
    for (const opt of sel.options) {
      if (opt.value === col) { sel.value = col; return; }
    }
  };

  if (chartType === "hist") {
    setIf(xSel, numericCols[0] || headers[0]);
    setIf(ySel, numericCols[1] || numericCols[0] || headers[0]);
    $("chartHint").textContent = "数値列を選ぶと分布が見えます。";
  } else if (chartType === "barCount") {
    setIf(xSel, stringCols[0] || headers[0]);
    setIf(ySel, numericCols[0] || headers[0]);
    $("chartHint").textContent = "カテゴリ列の件数を上位順に表示します（上位20）。";
  } else if (chartType === "scatter") {
    setIf(xSel, numericCols[0] || headers[0]);
    setIf(ySel, numericCols[1] || numericCols[0] || headers[0]);
    $("chartHint").textContent = "数値×数値の関係を散布図で表示します。";
  } else if (chartType === "lineTime") {
    setIf(xSel, dateCols[0] || headers[0]);
    setIf(ySel, numericCols[0] || headers[0]);
    $("chartHint").textContent = "日単位に集計して、Y列の平均を折れ線で表示します。";
  } else if (chartType === "corr") {
    setIf(xSel, numericCols[0] || headers[0]);
    setIf(ySel, numericCols[1] || numericCols[0] || headers[0]);
    $("chartHint").textContent = "数値列の相関（Pearson）をヒートマップ表示します（最大12列）。";
  } else if (chartType === "sankey") {
    $("chartHint").textContent = "source/target/value列があるシートで動作します。";
  } else {
    $("chartHint").textContent = "-";
  }
}

function drawChart() {
  const sheetName = $("chartsSheetSelect").value;
  const flags = getUIFlags();
  const rawRows = state.sheets[sheetName]?.rows || [];
  const headers = getVisibleHeaders(sheetName, flags);

  const chartType = $("chartType").value;
  const xCol = $("xCol").value;
  const yCol = $("yCol").value;

  const target = "chartArea";
  $("chartArea").innerHTML = ""; // clear

  if (rawRows.length === 0 || headers.length === 0) {
    $("chartArea").innerHTML = `<div class="text-muted small">データがありません</div>`;
    return;
  }

  if (chartType === "sankey") {
    if (headers.includes("source") && headers.includes("target") && headers.includes("value")) {
      drawSankey(target, sheetName);
    } else {
      $("chartArea").innerHTML = `<div class="text-muted small">必要な列（source/target/value）がありません</div>`;
    }
    return;
  }

  if (chartType === "corr") {
    // numeric columns
    const sample = (col) => rawRows.map(r => r[col]).filter(v => !isMissing(v)).slice(0, 200);
    const colTypes = {};
    for (const h of headers) colTypes[h] = detectColumnType(sample(h));
    const numericCols = headers.filter(h => colTypes[h] === "number").slice(0, 12);
    if (numericCols.length < 2) {
      $("chartArea").innerHTML = `<div class="text-muted small">数値列が不足しています</div>`;
      return;
    }
    const { cols, z } = buildCorrelationMatrix(rawRows, numericCols);
    Plotly.newPlot(target, [{
      type: "heatmap",
      x: cols,
      y: cols,
      z,
      zmin: -1,
      zmax: 1
    }], {
      margin: {l: 80, r: 10, t: 10, b: 80},
    }, {displayModeBar: true, responsive: true});
    return;
  }

  if (chartType === "hist") {
    const xs = rawRows.map(r => safeNumber(r[xCol])).filter(v => v !== null);
    if (xs.length === 0) {
      $("chartArea").innerHTML = `<div class="text-muted small">X列が数値として解釈できません</div>`;
      return;
    }
    Plotly.newPlot(target, [{
      type: "histogram",
      x: xs,
      nbinsx: Math.min(60, Math.max(10, Math.round(Math.sqrt(xs.length))))
    }], {
      margin: {l: 50, r: 10, t: 10, b: 40},
      xaxis: { title: xCol },
      yaxis: { title: "count" }
    }, {displayModeBar: true, responsive: true});
    return;
  }

  if (chartType === "barCount") {
    const vs = rawRows.map(r => r[xCol]).filter(v => !isMissing(v)).map(v => String(v));
    if (vs.length === 0) {
      $("chartArea").innerHTML = `<div class="text-muted small">X列が空です</div>`;
      return;
    }
    const top = freqTop(vs, 20);
    Plotly.newPlot(target, [{
      type: "bar",
      x: top.map(t => t.value),
      y: top.map(t => t.count),
    }], {
      margin: {l: 50, r: 10, t: 10, b: 120},
      xaxis: { tickangle: -45, title: xCol },
      yaxis: { title: "count" }
    }, {displayModeBar: true, responsive: true});
    return;
  }

  if (chartType === "scatter") {
    const x = [];
    const y = [];
    for (const r of rawRows) {
      const xi = safeNumber(r[xCol]);
      const yi = safeNumber(r[yCol]);
      if (xi === null || yi === null) continue;
      x.push(xi); y.push(yi);
    }
    if (x.length === 0) {
      $("chartArea").innerHTML = `<div class="text-muted small">X/Y列が数値として解釈できません</div>`;
      return;
    }
    Plotly.newPlot(target, [{
      type: "scatter",
      mode: "markers",
      x, y,
    }], {
      margin: {l: 50, r: 10, t: 10, b: 40},
      xaxis: { title: xCol },
      yaxis: { title: yCol },
    }, {displayModeBar: true, responsive: true});
    return;
  }

  if (chartType === "lineTime") {
    // group by day; y is average
    const m = new Map(); // day -> {sum, cnt}
    for (const r of rawRows) {
      const d = safeDate(r[xCol]);
      const yi = safeNumber(r[yCol]);
      if (!d || yi === null) continue;
      const day = toISODateString(d);
      if (!m.has(day)) m.set(day, { sum: 0, cnt: 0 });
      const obj = m.get(day);
      obj.sum += yi; obj.cnt += 1;
    }
    const days = Array.from(m.keys()).sort();
    if (days.length === 0) {
      $("chartArea").innerHTML = `<div class="text-muted small">X列が日時、Y列が数値として解釈できません</div>`;
      return;
    }
    const y = days.map(d => {
      const obj = m.get(d);
      return obj.cnt ? obj.sum / obj.cnt : 0;
    });
    Plotly.newPlot(target, [{
      type: "scatter",
      mode: "lines+markers",
      x: days,
      y,
    }], {
      margin: {l: 50, r: 10, t: 10, b: 40},
      xaxis: { title: xCol },
      yaxis: { title: `${yCol} (avg/day)` },
    }, {displayModeBar: true, responsive: true});
    return;
  }

  $("chartArea").innerHTML = `<div class="text-muted small">未対応のチャートタイプです</div>`;
}

function exportSheet(format) {
  const sheetName = $("exportSheetSelect").value;
  const flags = getUIFlags();
  const headers = getVisibleHeaders(sheetName, flags);
  const rows = getDisplayRows(sheetName, flags);

  const stamp = new Date().toISOString().replace(/[:.]/g,"-").slice(0,19);
  if (format === "csv") {
    const csv = toCSV(rows, headers);
    downloadText(`${sheetName}_${stamp}.csv`, csv, "text/csv");
  } else if (format === "json") {
    downloadText(`${sheetName}_${stamp}.json`, JSON.stringify(rows, null, 2), "application/json");
  }
}

async function loadWorkbookFromFile(file) {
  setStatus("読み込み中…", "muted");
  const buf = await readFileAsArrayBuffer(file);

  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  state.workbook = wb;
  state.sheets = {};
  state.sheetOrder = wb.SheetNames.slice();

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    // header:1 -> array-of-arrays (preserve column order)
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    const headersRaw = (aoa[0] || []).map((h, idx) => {
      if (h === null || h === undefined || String(h).trim() === "") return `col${idx+1}`;
      return String(h);
    });
    const rowsA = aoa.slice(1);

    // drop trailing fully-empty rows
    const rowsTrimmed = rowsA.filter(r => Array.isArray(r) && r.some(v => !isMissing(v)));

    const rows = rowsTrimmed.map(r => {
      const obj = {};
      for (let i = 0; i < headersRaw.length; i++) {
        obj[headersRaw[i]] = normalizeCellValue(r[i]);
      }
      return obj;
    });

    state.sheets[name] = { headers: headersRaw, rows };
  }

  populateSheetSelects();
  setStatus(`読込完了: ${file.name}（${state.sheetOrder.length} sheets）`, "success");

  // show contents
  ["overviewEmpty","diagEmpty","referralEmpty","sheetEmpty","profileEmpty","chartsEmpty","exportEmpty"].forEach(id => $(id).classList.add("d-none"));
  ["overviewContent","diagContent","referralContent","sheetContent","profileContent","chartsContent","exportContent"].forEach(id => $(id).classList.remove("d-none"));

  // set default sheet
  const preferred = state.sheets["diagnosis"] ? "diagnosis" : state.sheetOrder[0];
  setCurrentSheet(preferred);

  renderOverview();
  renderDiagnosisInit();
  renderReferralDeepDive();
}



/* -----------------------------
 * Diagnosis analysis (specialized)
 * ----------------------------- */

function parseAgeNumber(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "string") {
    const s = v.trim();
    if (s === "" || s.toLowerCase() === "nan") return null;

    // "26+"
    let m = s.match(/^(\d+)\s*\+\s*$/);
    if (m) return Number(m[1]);

    // "23-25" / "23〜25" / "23~25" etc
    m = s.match(/^(\d+)\s*[-〜~–—]\s*(\d+)\s*$/);
    if (m) return (Number(m[1]) + Number(m[2])) / 2;

    // plain numeric
    const x = Number(s);
    if (Number.isFinite(x)) return x;

    // fallback: first number
    m = s.match(/(\d+)/);
    if (m) return Number(m[1]);
  }
  return null;
}

function normalizeGender(v) {
  const s = String(v || "").trim().toLowerCase();
  if (s === "female" || s === "male") return s;
  return "unknown";
}

function parseInterested(v) {
  // interested is expected to be 1 or null; handle string/boolean defensively
  if (v === true) return 1;
  if (v === false) return 0;
  const n = safeNumber(v);
  if (n === 1) return 1;
  return 0;
}

function parseAnswersFromRawJson(raw) {
  if (isMissing(raw)) return null;
  if (typeof raw === "object") {
    const ans = raw.answers;
    if (ans && typeof ans === "object") return ans;
    return null;
  }
  if (typeof raw !== "string") return null;
  try {
    const obj = JSON.parse(raw);
    const ans = obj?.answers;
    if (ans && typeof ans === "object") return ans;
    return null;
  } catch (e) {
    return null;
  }
}

function sortQuestionKeys(keys) {
  return keys.slice().sort((a, b) => {
    const ma = String(a).match(/^([A-Za-z]+)(\d+)$/);
    const mb = String(b).match(/^([A-Za-z]+)(\d+)$/);
    if (ma && mb) {
      if (ma[1] !== mb[1]) return ma[1].localeCompare(mb[1]);
      return Number(ma[2]) - Number(mb[2]);
    }
    return String(a).localeCompare(String(b));
  });
}

function getAgeBinLabel(age) {
  if (age === null || age === undefined || !Number.isFinite(age)) return "unknown";
  if (age < 16) return "13-15";
  if (age < 19) return "16-18";
  if (age < 23) return "19-22";
  return "23+";
}

function dominantAxisLabel(r) {
  const a = r.axisA, b = r.axisB, c = r.axisC, d = r.axisD;
  const arr = [
    { k: "A", v: a },
    { k: "B", v: b },
    { k: "C", v: c },
    { k: "D", v: d },
  ].filter(o => o.v !== null && o.v !== undefined && Number.isFinite(o.v));
  if (arr.length === 0) return null;
  arr.sort((x, y) => y.v - x.v);
  return arr[0].k;
}

function prepareDiagnosisCache() {
  if (!state.sheets["diagnosis"]) {
    state.diag = null;
    return;
  }

  const rawRows = state.sheets["diagnosis"].rows || [];
  const rows = [];

  const typeCount = new Map();
  const qKeySet = new Set();
  let parseOk = 0;

  let minDay = null, maxDay = null;
  let minAge = null, maxAge = null;

  for (const r of rawRows) {
    const d = safeDate(r["createdAt"]);
    const day = d ? toISODateString(d) : null;

    if (day) {
      if (!minDay || day < minDay) minDay = day;
      if (!maxDay || day > maxDay) maxDay = day;
    }

    const type = isMissing(r["type"]) ? "(missing)" : String(r["type"]);
    typeCount.set(type, (typeCount.get(type) || 0) + 1);

    const ageRaw = r["age"];
    const ageNum = parseAgeNumber(ageRaw);
    if (ageNum !== null) {
      minAge = (minAge === null) ? ageNum : Math.min(minAge, ageNum);
      maxAge = (maxAge === null) ? ageNum : Math.max(maxAge, ageNum);
    }

    const answers = parseAnswersFromRawJson(r["raw_json"]);
    if (answers) {
      parseOk++;
      for (const k of Object.keys(answers)) qKeySet.add(k);
    }

    rows.push({
      createdAt: d,
      createdDay: day,
      email: r["email"],
      gender: normalizeGender(r["gender"]),
      type,
      age_raw: ageRaw,
      age_num: ageNum,
      axisA: safeNumber(r["axisA"]),
      axisB: safeNumber(r["axisB"]),
      axisC: safeNumber(r["axisC"]),
      axisD: safeNumber(r["axisD"]),
      interested: parseInterested(r["interested"]),
      raw_json: r["raw_json"],
      answers: answers || null,
    });
  }

  const typesSorted = Array.from(typeCount.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([t, c]) => ({ type: t, count: c }));

  const qKeys = sortQuestionKeys(Array.from(qKeySet));

  state.diag = {
    rows,
    typeCount,
    typesSorted,
    qKeys,
    parseOk,
    total: rows.length,
    minDay,
    maxDay,
    minAge,
    maxAge,
    defaults: {
      dateStart: minDay || "",
      dateEnd: maxDay || "",
      gender: "all",
      type: "all",
      ageMin: (minAge !== null) ? String(Math.floor(minAge)) : "",
      ageMax: (maxAge !== null) ? String(Math.ceil(maxAge)) : "",
      includeMissingAge: true,
      axis: "axisA",
      axisGroupBy: "type",
      axisTopN: 10,
      qKey: qKeys[0] || "",
    }
  };
}

function setupDiagnosisControls() {
  if (!state.diag) return;

  // type select
  const typeSel = $("diagType");
  if (typeSel) {
    typeSel.innerHTML = "";
    const optAll = document.createElement("option");
    optAll.value = "all"; optAll.textContent = "all";
    typeSel.appendChild(optAll);
    for (const obj of state.diag.typesSorted) {
      const opt = document.createElement("option");
      opt.value = obj.type;
      opt.textContent = `${obj.type} (${obj.count})`;
      typeSel.appendChild(opt);
    }
  }

  // question key select
  const qSel = $("diagQKey");
  if (qSel) {
    qSel.innerHTML = "";
    if (state.diag.qKeys.length === 0) {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "(answers が見つかりません)";
      qSel.appendChild(opt);
    } else {
      for (const k of state.diag.qKeys) {
        const opt = document.createElement("option");
        opt.value = k;
        opt.textContent = k;
        qSel.appendChild(opt);
      }
    }
  }

  resetDiagnosisFiltersToDefaults(false);
}

function resetDiagnosisFiltersToDefaults(renderAfter = true) {
  if (!state.diag) return;
  const d = state.diag.defaults;

  if ($("diagDateStart")) $("diagDateStart").value = d.dateStart || "";
  if ($("diagDateEnd")) $("diagDateEnd").value = d.dateEnd || "";
  if ($("diagGender")) $("diagGender").value = d.gender || "all";
  if ($("diagType")) $("diagType").value = d.type || "all";
  if ($("diagAgeMin")) $("diagAgeMin").value = d.ageMin || "";
  if ($("diagAgeMax")) $("diagAgeMax").value = d.ageMax || "";
  if ($("diagIncludeMissingAge")) $("diagIncludeMissingAge").checked = !!d.includeMissingAge;

  if ($("diagAxisSelect")) $("diagAxisSelect").value = d.axis || "axisA";
  if ($("diagAxisGroupBy")) $("diagAxisGroupBy").value = d.axisGroupBy || "type";
  if ($("diagAxisTopN")) $("diagAxisTopN").value = String(d.axisTopN || 10);

  if ($("diagQKey")) $("diagQKey").value = d.qKey || "";

  if (renderAfter) renderDiagnosisAnalysis();
}

function getDiagnosisFiltersFromUI() {
  const ageMin = safeNumber($("diagAgeMin")?.value);
  const ageMax = safeNumber($("diagAgeMax")?.value);

  const topNRaw = safeNumber($("diagAxisTopN")?.value);
  const axisTopN = Math.max(3, Math.min(30, Math.round(topNRaw || 10)));

  return {
    dateStart: $("diagDateStart")?.value || null,
    dateEnd: $("diagDateEnd")?.value || null,
    gender: $("diagGender")?.value || "all",
    type: $("diagType")?.value || "all",
    ageMin: ageMin === null ? null : ageMin,
    ageMax: ageMax === null ? null : ageMax,
    includeMissingAge: $("diagIncludeMissingAge")?.checked ?? true,
    axis: $("diagAxisSelect")?.value || "axisA",
    axisGroupBy: $("diagAxisGroupBy")?.value || "type",
    axisTopN,
    qKey: $("diagQKey")?.value || "",
  };
}

function filterDiagnosisRows(rows, f) {
  return rows.filter(r => {
    if ((f.dateStart || f.dateEnd) && !r.createdDay) return false;
    if (f.dateStart && r.createdDay < f.dateStart) return false;
    if (f.dateEnd && r.createdDay > f.dateEnd) return false;

    if (f.gender !== "all" && r.gender !== f.gender) return false;
    if (f.type !== "all" && r.type !== f.type) return false;

    const age = r.age_num;
    if (age === null || age === undefined || !Number.isFinite(age)) {
      if (!f.includeMissingAge && (f.ageMin !== null || f.ageMax !== null)) return false;
    } else {
      if (f.ageMin !== null && age < f.ageMin) return false;
      if (f.ageMax !== null && age > f.ageMax) return false;
    }
    return true;
  });
}

function meanOf(nums) {
  const xs = nums.filter(v => v !== null && v !== undefined && Number.isFinite(v));
  if (xs.length === 0) return null;
  return xs.reduce((a, b) => a + b, 0) / xs.length;
}

function medianOf(nums) {
  const xs = nums.filter(v => v !== null && v !== undefined && Number.isFinite(v)).sort((a, b) => a - b);
  if (xs.length === 0) return null;
  return median(xs);
}

function renderDiagnosisInit() {
  // Called after workbook load
  const hasDiagnosis = !!state.sheets["diagnosis"];
  if (!hasDiagnosis) {
    if ($("diagNoSheet")) $("diagNoSheet").classList.remove("d-none");
    if ($("diagContent")) $("diagContent").classList.add("d-none");
    return;
  }
  if ($("diagNoSheet")) $("diagNoSheet").classList.add("d-none");
  if ($("diagContent")) $("diagContent").classList.remove("d-none");

  prepareDiagnosisCache();
  setupDiagnosisControls();
  renderDiagnosisAnalysis();
}

function renderDiagnosisAnalysis() {
  if (!state.diag) return;

  const f = getDiagnosisFiltersFromUI();
  const rows = filterDiagnosisRows(state.diag.rows, f);

  if ($("diagN")) $("diagN").textContent = `N=${rows.length.toLocaleString()}`;

  renderDiagnosisKPIs(rows);
  renderDiagnosisTypeCharts(rows);
  renderDiagnosisTypeTable(rows);
  renderDiagnosisAxisBox(rows, f);
  renderDiagnosisCorr(rows);
  renderDiagnosisInterestedSegments(rows);
  renderDiagnosisDominantAxis(rows);
  renderDiagnosisQuestion(rows, f.qKey);
  renderDiagnosisQuality(rows);
}

function renderDiagnosisKPIs(rows) {
  const n = rows.length;
  if (n === 0) {
    ["diagKpiRows","diagKpiInterestedRate","diagKpiUniqueEmail","diagKpiAgeMedian","diagKpiFemaleRate","diagKpiJsonParseRate"]
      .forEach(id => { if ($(id)) $(id).textContent = "-"; });
    return;
  }

  const interested = rows.reduce((s, r) => s + (r.interested || 0), 0);
  const rate = interested / n;

  const emails = new Set(rows.map(r => r.email).filter(v => !isMissing(v)).map(v => String(v)));
  const ages = rows.map(r => r.age_num);
  const ageMed = medianOf(ages);

  const female = rows.filter(r => r.gender === "female").length;
  const femaleRate = female / n;

  const parsed = rows.filter(r => r.answers).length;
  const parseRate = parsed / n;

  if ($("diagKpiRows")) $("diagKpiRows").textContent = n.toLocaleString();
  if ($("diagKpiInterestedRate")) $("diagKpiInterestedRate").textContent = (rate * 100).toFixed(1) + "%";
  if ($("diagKpiUniqueEmail")) $("diagKpiUniqueEmail").textContent = emails.size.toLocaleString();
  if ($("diagKpiAgeMedian")) $("diagKpiAgeMedian").textContent = (ageMed === null ? "-" : ageMed.toFixed(1).replace(/\.0$/, ""));
  if ($("diagKpiFemaleRate")) $("diagKpiFemaleRate").textContent = (femaleRate * 100).toFixed(1) + "%";
  if ($("diagKpiJsonParseRate")) $("diagKpiJsonParseRate").textContent = (parseRate * 100).toFixed(1) + "%";
}

function renderDiagnosisTypeCharts(rows) {
  const typeM = new Map();
  const intM = new Map(); // type -> interested sum
  for (const r of rows) {
    const t = r.type || "(missing)";
    typeM.set(t, (typeM.get(t) || 0) + 1);
    intM.set(t, (intM.get(t) || 0) + (r.interested || 0));
  }

  const arr = Array.from(typeM.entries()).map(([t, c]) => ({ type: t, count: c, interested: intM.get(t) || 0 }));
  arr.sort((a, b) => b.count - a.count);

  // Count chart (top 20 + Other)
  const topK = 20;
  const top = arr.slice(0, topK);
  const rest = arr.slice(topK);
  const otherCount = rest.reduce((s, o) => s + o.count, 0);
  const x1 = top.map(o => o.type);
  const y1 = top.map(o => o.count);
  if (otherCount > 0) { x1.push("Other"); y1.push(otherCount); }

  Plotly.newPlot("plotDiagTypeCount", [{
    type: "bar",
    x: x1,
    y: y1,
  }], {
    margin: {l: 50, r: 10, t: 10, b: 120},
    xaxis: { tickangle: -45, title: "type" },
    yaxis: { title: "count" }
  }, {displayModeBar: true, responsive: true});

  // Interested rate chart (top 20 by count)
  const x2 = top.map(o => o.type);
  const y2 = top.map(o => o.count ? (o.interested / o.count) : 0);

  Plotly.newPlot("plotDiagInterestedByType", [{
    type: "bar",
    x: x2,
    y: y2,
    text: y2.map(v => (v * 100).toFixed(1) + "%"),
    textposition: "auto",
  }], {
    margin: {l: 50, r: 10, t: 10, b: 120},
    xaxis: { tickangle: -45, title: "type" },
    yaxis: { title: "Interested rate", tickformat: ".0%", range: [0, 1] }
  }, {displayModeBar: true, responsive: true});
}

function renderDiagnosisTypeTable(rows) {
  const typeM = new Map();
  for (const r of rows) {
    const t = r.type || "(missing)";
    if (!typeM.has(t)) typeM.set(t, { type: t, n: 0, interested: 0, axisA: [], axisB: [], axisC: [], axisD: [], ages: [] });
    const o = typeM.get(t);
    o.n++;
    o.interested += (r.interested || 0);
    if (Number.isFinite(r.axisA)) o.axisA.push(r.axisA);
    if (Number.isFinite(r.axisB)) o.axisB.push(r.axisB);
    if (Number.isFinite(r.axisC)) o.axisC.push(r.axisC);
    if (Number.isFinite(r.axisD)) o.axisD.push(r.axisD);
    if (Number.isFinite(r.age_num)) o.ages.push(r.age_num);
  }

  const arr = Array.from(typeM.values()).map(o => ({
    type: o.type,
    n: o.n,
    interested: o.interested,
    rate: o.n ? o.interested / o.n : 0,
    axisA: meanOf(o.axisA),
    axisB: meanOf(o.axisB),
    axisC: meanOf(o.axisC),
    axisD: meanOf(o.axisD),
    age_mean: meanOf(o.ages),
  })).sort((a, b) => b.n - a.n);

  const tbody = $("diagTypeTable")?.querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  const show = arr.slice(0, 50);
  for (const o of show) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="codeLike">${o.type}</td>
      <td class="text-end">${o.n.toLocaleString()}</td>
      <td class="text-end">${o.interested.toLocaleString()}</td>
      <td class="text-end">${(o.rate * 100).toFixed(1)}%</td>
      <td class="text-end">${(o.axisA === null ? "" : o.axisA.toFixed(1))}</td>
      <td class="text-end">${(o.axisB === null ? "" : o.axisB.toFixed(1))}</td>
      <td class="text-end">${(o.axisC === null ? "" : o.axisC.toFixed(1))}</td>
      <td class="text-end">${(o.axisD === null ? "" : o.axisD.toFixed(1))}</td>
      <td class="text-end">${(o.age_mean === null ? "" : o.age_mean.toFixed(1).replace(/\.0$/, ""))}</td>
    `;
    tbody.appendChild(tr);
  }

  // DataTables init/destroy for diagnosis table
  const table = $("diagTypeTable");
  if (state.diagTypeDT) {
    try { state.diagTypeDT.destroy(); } catch(e) {}
    state.diagTypeDT = null;
  }
  state.diagTypeDT = new DataTable(table, {
    paging: true,
    pageLength: 25,
    lengthMenu: [ [25, 50, 100], [25, 50, 100] ],
    searching: true,
    info: true,
    responsive: true,
    order: [[1, "desc"]],
  });
}

function renderDiagnosisAxisBox(rows, f) {
  const axis = f.axis;
  const groupBy = f.axisGroupBy;

  if (!["axisA","axisB","axisC","axisD"].includes(axis)) return;

  let groups = new Map();

  if (groupBy === "gender") {
    const order = ["female","male","unknown"];
    for (const g of order) groups.set(g, []);
    for (const r of rows) {
      const v = r[axis];
      if (!Number.isFinite(v)) continue;
      const g = r.gender || "unknown";
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(v);
    }
  } else if (groupBy === "ageBin") {
    const order = ["13-15","16-18","19-22","23+","unknown"];
    for (const b of order) groups.set(b, []);
    for (const r of rows) {
      const v = r[axis];
      if (!Number.isFinite(v)) continue;
      const b = getAgeBinLabel(r.age_num);
      if (!groups.has(b)) groups.set(b, []);
      groups.get(b).push(v);
    }
  } else {
    // groupBy type (top N)
    const cnt = new Map();
    for (const r of rows) cnt.set(r.type, (cnt.get(r.type) || 0) + 1);
    const topTypes = Array.from(cnt.entries()).sort((a, b) => b[1] - a[1]).slice(0, f.axisTopN).map(x => x[0]);
    for (const t of topTypes) groups.set(t, []);
    for (const r of rows) {
      if (!groups.has(r.type)) continue;
      const v = r[axis];
      if (!Number.isFinite(v)) continue;
      groups.get(r.type).push(v);
    }
  }

  // build traces
  const traces = [];
  for (const [k, vals] of groups.entries()) {
    if (!vals || vals.length === 0) continue;
    traces.push({ type: "box", name: k, y: vals });
  }
  if (traces.length === 0) {
    $("plotDiagAxisBox").innerHTML = `<div class="text-muted small">表示できるデータがありません</div>`;
    return;
  }

  Plotly.newPlot("plotDiagAxisBox", traces, {
    margin: {l: 50, r: 10, t: 10, b: 120},
    yaxis: { title: axis, range: [0, 100] },
    boxmode: "group"
  }, {displayModeBar: true, responsive: true});
}

function renderDiagnosisCorr(rows) {
  const cols = ["axisA","axisB","axisC","axisD","age_num","interested"];
  const { cols: c, z } = buildCorrelationMatrix(rows, cols);

  // If everything is null, show message
  const any = z.flat().some(v => v !== null);
  if (!any) {
    $("plotDiagCorr").innerHTML = `<div class="text-muted small">相関を計算できるデータが不足しています</div>`;
    return;
  }

  Plotly.newPlot("plotDiagCorr", [{
    type: "heatmap",
    x: c,
    y: c,
    z,
    zmin: -1,
    zmax: 1
  }], {
    margin: {l: 80, r: 10, t: 10, b: 80},
  }, {displayModeBar: true, responsive: true});
}

function renderDiagnosisInterestedSegments(rows) {
  // by age bin
  const bins = ["13-15","16-18","19-22","23+","unknown"];
  const m = new Map(bins.map(b => [b, { n: 0, interested: 0 }]));
  for (const r of rows) {
    const b = getAgeBinLabel(r.age_num);
    if (!m.has(b)) m.set(b, { n: 0, interested: 0 });
    const o = m.get(b);
    o.n++;
    o.interested += (r.interested || 0);
  }
  const xA = bins;
  const yA = bins.map(b => {
    const o = m.get(b);
    return o && o.n ? (o.interested / o.n) : 0;
  });

  Plotly.newPlot("plotDiagInterestedByAge", [{
    type: "bar",
    x: xA,
    y: yA,
    text: yA.map(v => (v * 100).toFixed(1) + "%"),
    textposition: "auto",
  }], {
    margin: {l: 50, r: 10, t: 10, b: 80},
    yaxis: { title: "Interested rate", tickformat: ".0%", range: [0, 1] },
    xaxis: { title: "age bin" },
  }, {displayModeBar: true, responsive: true});

  // by gender
  const genders = ["female","male","unknown"];
  const mg = new Map(genders.map(g => [g, { n: 0, interested: 0 }]));
  for (const r of rows) {
    const g = r.gender || "unknown";
    if (!mg.has(g)) mg.set(g, { n: 0, interested: 0 });
    const o = mg.get(g);
    o.n++;
    o.interested += (r.interested || 0);
  }
  const xG = genders;
  const yG = genders.map(g => {
    const o = mg.get(g);
    return o && o.n ? (o.interested / o.n) : 0;
  });

  Plotly.newPlot("plotDiagInterestedByGender", [{
    type: "bar",
    x: xG,
    y: yG,
    text: yG.map(v => (v * 100).toFixed(1) + "%"),
    textposition: "auto",
  }], {
    margin: {l: 50, r: 10, t: 10, b: 80},
    yaxis: { title: "Interested rate", tickformat: ".0%", range: [0, 1] },
    xaxis: { title: "gender" },
  }, {displayModeBar: true, responsive: true});
}

function renderDiagnosisDominantAxis(rows) {
  const cats = ["A","B","C","D"];
  const m = new Map(cats.map(c => [c, { n: 0, interested: 0 }]));
  for (const r of rows) {
    const k = dominantAxisLabel(r);
    if (!k) continue;
    const o = m.get(k);
    o.n++;
    o.interested += (r.interested || 0);
  }
  const x = cats;
  const y = cats.map(c => m.get(c).n);
  const rate = cats.map(c => {
    const o = m.get(c);
    return o.n ? (o.interested / o.n) : 0;
  });

  Plotly.newPlot("plotDiagDominantAxis", [{
    type: "bar",
    x,
    y,
    text: rate.map(v => "rate " + (v * 100).toFixed(1) + "%"),
    textposition: "auto",
  }], {
    margin: {l: 50, r: 10, t: 10, b: 60},
    xaxis: { title: "dominant axis" },
    yaxis: { title: "count" }
  }, {displayModeBar: true, responsive: true});
}

function renderDiagnosisQuestion(rows, qKey) {
  const histDiv = $("plotDiagQHist");
  const meanDiv = $("plotDiagQMeanByType");
  if (!histDiv || !meanDiv) return;

  if (!qKey) {
    histDiv.innerHTML = `<div class="text-muted small">answers のキーが見つかりません</div>`;
    meanDiv.innerHTML = `<div class="text-muted small">answers のキーが見つかりません</div>`;
    ["diagQMeanInterested","diagQMeanNotInterested","diagQMeanDiff"].forEach(id => { if ($(id)) $(id).textContent = "-"; });
    return;
  }

  const vals = [];
  const byType = new Map();
  let sum1 = 0, n1 = 0, sum0 = 0, n0 = 0;

  for (const r of rows) {
    if (!r.answers) continue;
    const v = safeNumber(r.answers[qKey]);
    if (v === null) continue;
    vals.push(v);

    const t = r.type || "(missing)";
    if (!byType.has(t)) byType.set(t, { sum: 0, n: 0 });
    const o = byType.get(t);
    o.sum += v; o.n++;

    if (r.interested === 1) { sum1 += v; n1++; }
    else { sum0 += v; n0++; }
  }

  if (vals.length === 0) {
    histDiv.innerHTML = `<div class="text-muted small">フィルタ条件では ${qKey} の回答が取得できません</div>`;
    meanDiv.innerHTML = `<div class="text-muted small">フィルタ条件では ${qKey} の回答が取得できません</div>`;
    ["diagQMeanInterested","diagQMeanNotInterested","diagQMeanDiff"].forEach(id => { if ($(id)) $(id).textContent = "-"; });
    return;
  }

  // Histogram
  Plotly.newPlot("plotDiagQHist", [{
    type: "histogram",
    x: vals,
    xbins: { start: 0.5, end: 7.5, size: 1 },
  }], {
    margin: {l: 50, r: 10, t: 10, b: 60},
    xaxis: { title: qKey, dtick: 1 },
    yaxis: { title: "count" }
  }, {displayModeBar: true, responsive: true});

  // Mean by type (top 15 by n)
  const arr = Array.from(byType.entries())
    .map(([t, o]) => ({ type: t, n: o.n, mean: o.n ? o.sum / o.n : 0 }))
    .sort((a, b) => b.n - a.n)
    .slice(0, 15);

  Plotly.newPlot("plotDiagQMeanByType", [{
    type: "bar",
    x: arr.map(o => o.type),
    y: arr.map(o => o.mean),
    text: arr.map(o => o.mean.toFixed(2)),
    textposition: "auto",
  }], {
    margin: {l: 50, r: 10, t: 10, b: 120},
    xaxis: { tickangle: -45, title: "type" },
    yaxis: { title: `mean(${qKey})`, range: [0, 7.5] },
  }, {displayModeBar: true, responsive: true});

  const m1 = n1 ? (sum1 / n1) : null;
  const m0 = n0 ? (sum0 / n0) : null;
  const diff = (m1 !== null && m0 !== null) ? (m1 - m0) : null;

  if ($("diagQMeanInterested")) $("diagQMeanInterested").textContent = (m1 === null ? "-" : m1.toFixed(2));
  if ($("diagQMeanNotInterested")) $("diagQMeanNotInterested").textContent = (m0 === null ? "-" : m0.toFixed(2));
  if ($("diagQMeanDiff")) $("diagQMeanDiff").textContent = (diff === null ? "-" : diff.toFixed(2));
}

function renderDiagnosisQuality(rows) {
  const n = rows.length;
  if (n === 0) {
    ["diagQualityDupEmail","diagQualityBadAge","diagQualityAxisOut","diagQualityNoDate"].forEach(id => { if ($(id)) $(id).textContent = "-"; });
    return;
  }

  // duplicate emails
  const em = new Map();
  for (const r of rows) {
    if (isMissing(r.email)) continue;
    const e = String(r.email);
    em.set(e, (em.get(e) || 0) + 1);
  }
  let dupEmails = 0;
  let extraRows = 0;
  let maxDup = 0;
  for (const c of em.values()) {
    if (c > 1) {
      dupEmails++;
      extraRows += (c - 1);
      maxDup = Math.max(maxDup, c);
    }
  }

  // bad age parse: age_raw exists but age_num is null
  let badAge = 0;
  for (const r of rows) {
    if (!isMissing(r.age_raw) && (r.age_num === null || r.age_num === undefined || !Number.isFinite(r.age_num))) badAge++;
  }

  // axis out of range
  let axisOut = 0;
  for (const r of rows) {
    const xs = [r.axisA, r.axisB, r.axisC, r.axisD].filter(v => v !== null && v !== undefined && Number.isFinite(v));
    if (xs.some(v => v < 0 || v > 100)) axisOut++;
  }

  // no createdAt
  const noDate = rows.filter(r => !r.createdDay).length;

  if ($("diagQualityDupEmail")) $("diagQualityDupEmail").textContent = `${dupEmails.toLocaleString()}（余剰${extraRows.toLocaleString()} / 最大${maxDup}）`;
  if ($("diagQualityBadAge")) $("diagQualityBadAge").textContent = badAge.toLocaleString();
  if ($("diagQualityAxisOut")) $("diagQualityAxisOut").textContent = axisOut.toLocaleString();
  if ($("diagQualityNoDate")) $("diagQualityNoDate").textContent = noDate.toLocaleString();
}

function wireEvents() {
  $("fileInput").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      await loadWorkbookFromFile(file);
    } catch (err) {
      console.error(err);
      setStatus("読み込みに失敗しました（コンソールも確認してください）", "danger");
      alert("読み込みに失敗しました。ファイル形式や破損を確認してください。");
    }
  });

  $("maskPII").addEventListener("change", () => {
    if (!state.workbook) return;
    renderSheetTable();
    renderProfile();
    renderReferralDeepDive();
  });
  $("showJsonCols").addEventListener("change", () => {
    if (!state.workbook) return;
    renderOverview();
    populateSheetSelects();
    setCurrentSheet(state.currentSheet || state.sheetOrder[0]);
    renderReferralDeepDive();
  });
  $("showUnnamedCols").addEventListener("change", () => {
    if (!state.workbook) return;
    renderOverview();
    populateSheetSelects();
    setCurrentSheet(state.currentSheet || state.sheetOrder[0]);
    renderReferralDeepDive();
  });

  $("sheetSelect").addEventListener("change", () => {
    renderSheetTable();
  });

  $("btnRefreshView").addEventListener("click", () => {
    renderSheetTable();
  });

  $("btnGoProfile").addEventListener("click", () => {
    const tabEl = document.querySelector('#tab-profile');
    const tab = new bootstrap.Tab(tabEl);
    tab.show();
    $("profileSheetSelect").value = $("sheetSelect").value;
    renderProfile();
  });

  $("btnGoCharts").addEventListener("click", () => {
    const tabEl = document.querySelector('#tab-charts');
    const tab = new bootstrap.Tab(tabEl);
    tab.show();
    $("chartsSheetSelect").value = $("sheetSelect").value;
    syncChartColumns();
  });

  $("profileSheetSelect").addEventListener("change", () => renderProfile());
  $("btnProfileRecalc").addEventListener("click", () => renderProfile());

  $("chartsSheetSelect").addEventListener("change", () => syncChartColumns());
  $("chartType").addEventListener("change", () => autoPickColumns());
  $("btnAutoPick").addEventListener("click", () => autoPickColumns());
  $("btnDrawChart").addEventListener("click", () => drawChart());

  $("exportSheetSelect").addEventListener("change", () => {});
  $("btnExportCsv").addEventListener("click", () => exportSheet("csv"));
  $("btnExportJson").addEventListener("click", () => exportSheet("json"));
  // Referral deep dive
  const refApply = $("btnRefApply");
  if (refApply) refApply.addEventListener("click", () => {
    if (!state.workbook) return;
    renderReferralDeepDive();
  });

  const refReset = $("btnRefReset");
  if (refReset) refReset.addEventListener("click", () => {
    if (!state.workbook) return;
    // reset to full bounds
    const eventsAll = getReferralEventsAll();
    const b = computeEventDateBounds(eventsAll);
    if (b.min && b.max) {
      const s = $("refStartDate");
      const e = $("refEndDate");
      if (s) s.value = toISODateString(b.min);
      if (e) e.value = toISODateString(b.max);
    }
    renderReferralDeepDive();
  });

  const refSel = $("referrerSelect");
  if (refSel) refSel.addEventListener("change", () => {
    state.referralSelectedReferrer = refSel.value;
    renderReferrerDetail();
  });

  const refDetailRefresh = $("btnRefDetailRefresh");
  if (refDetailRefresh) refDetailRefresh.addEventListener("click", () => {
    renderReferrerDetail();
  });

  const sankeyKind = $("refSankeyKind");
  if (sankeyKind) sankeyKind.addEventListener("change", () => drawReferralSankey());

  const sankeyVariant = $("refSankeyVariant");
  if (sankeyVariant) sankeyVariant.addEventListener("change", () => drawReferralSankey());

  const sankeyMin = $("refSankeyMinValue");
  if (sankeyMin) {
    sankeyMin.addEventListener("change", () => drawReferralSankey());
    sankeyMin.addEventListener("keyup", (e) => {
      if (e.key === "Enter") drawReferralSankey();
    });
  }
  // Leaderboard row click -> detail select (event delegation)
  const refTable = $("referrerTable");
  if (refTable) refTable.addEventListener("click", (e) => {
    const tr = e.target.closest("tr");
    if (!tr) return;
    const rid = tr.getAttribute("data-referrer-id");
    if (!rid) return;
    state.referralSelectedReferrer = rid;
    const sel = $("referrerSelect");
    if (sel) sel.value = rid;
    renderReferrerDetail();
  });
}

wireEvents();
