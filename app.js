
/* eslint-disable no-console */
(function () {
  'use strict';

  // -----------------------------
  // State
  // -----------------------------
  const state = {
    file: null,
    raw: {
      diagnosis: [],
      referral_events: []
    },
    norm: {
      diagnosis: [],
      referral_events: []
    },
    derived: {
      diagnosisUsers: [],          // user-level summaries (latest record, favorite flags, etc.)
      diagnosisUserIndex: new Map(), // emailLower -> user summary
      refEvents: [],               // normalized referral events
      refDaily: [],                // daily aggregates
      referrerMeta: new Map(),     // referrerId -> meta
      userMeta: new Map(),         // userId -> meta (from completes)
      edges: {
        visits: [],
        completes: []
      },
      completeEmailToReferrer: new Map(), // emailLower -> {referrerId, ts}
    },
    ui: {
      maskPII: true,
      hideUnnamed: true
    },
    dt: {
      tableDiagnosis: null,
      tableFavorites: null,
      tableRefEvents: null,
      tableReferrers: null,
      tableRefEdges: null,
    },
    selection: {
      selectedReferrerId: null
    }
  };

  // -----------------------------
  // DOM helpers
  // -----------------------------
  const el = (id) => document.getElementById(id);

  function setText(id, text) {
    const node = el(id);
    if (!node) return;
    node.textContent = (text ?? '').toString();
  }

  function show(id) {
    const node = el(id);
    if (!node) return;
    node.classList.remove('d-none');
  }

  function hide(id) {
    const node = el(id);
    if (!node) return;
    node.classList.add('d-none');
  }

  function escapeHtml(str) {
    return (str ?? '').toString()
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }

  // -----------------------------
  // Utility: dates, numbers, parsing
  // -----------------------------
  function isValidDate(d) {
    return d instanceof Date && !Number.isNaN(d.getTime());
  }

  function parseISODate(s) {
    if (!s) return null;
    const d = new Date(s);
    if (isValidDate(d)) return d;
    return null;
  }

  function toISODateString(d) {
    if (!d) return '';
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  function clamp01(x) {
    if (!Number.isFinite(x)) return 0;
    return Math.max(0, Math.min(1, x));
  }

  function safeNumber(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }

  function parseAgeToNumber(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    const s = String(v).trim();
    if (!s) return null;

    // "23-25" => 24
    const mRange = s.match(/^(\d{1,3})\s*-\s*(\d{1,3})$/);
    if (mRange) {
      const a = Number(mRange[1]);
      const b = Number(mRange[2]);
      if (Number.isFinite(a) && Number.isFinite(b)) return (a + b) / 2;
    }

    // "26+" => 26
    const mPlus = s.match(/^(\d{1,3})\s*\+$/);
    if (mPlus) {
      const a = Number(mPlus[1]);
      if (Number.isFinite(a)) return a;
    }

    // plain integer
    const mInt = s.match(/^(\d{1,3})$/);
    if (mInt) {
      const a = Number(mInt[1]);
      if (Number.isFinite(a)) return a;
    }

    return null;
  }

  function parseJsonSafe(s) {
    if (s === null || s === undefined) return null;
    if (typeof s === 'object') return s;
    const str = String(s);
    if (!str) return null;
    try {
      return JSON.parse(str);
    } catch (e) {
      return null;
    }
  }

  function shortHash(str) {
    // deterministic, simple hash -> base36
    const s = (str ?? '').toString();
    let h = 5381;
    for (let i = 0; i < s.length; i++) {
      h = ((h << 5) + h) + s.charCodeAt(i); // djb2
      h = h >>> 0;
    }
    return h.toString(36).padStart(6, '0');
  }

  function maskId(kind, raw) {
    if (!raw) return '';
    return `${kind}_${shortHash(raw)}`;
  }

  function normalizeGender(g) {
    const s = (g ?? '').toString().trim().toLowerCase();
    if (s === 'female' || s === 'f') return 'female';
    if (s === 'male' || s === 'm') return 'male';
    if (!s) return 'unknown';
    return s; // keep other labels
  }

  function normalizeType(t) {
    const s = (t ?? '').toString().trim();
    return s || '(unknown)';
  }

  function normalizeEmail(e) {
    const s = (e ?? '').toString().trim().toLowerCase();
    return s || null;
  }

  function formatPct(x) {
    if (!Number.isFinite(x)) return '—';
    return `${(x * 100).toFixed(1)}%`;
  }

  function mean(arr) {
    const xs = arr.filter(x => Number.isFinite(x));
    if (!xs.length) return null;
    const s = xs.reduce((a, b) => a + b, 0);
    return s / xs.length;
  }

  function median(arr) {
    const xs = arr.filter(x => Number.isFinite(x)).sort((a, b) => a - b);
    if (!xs.length) return null;
    const mid = Math.floor(xs.length / 2);
    if (xs.length % 2 === 1) return xs[mid];
    return (xs[mid - 1] + xs[mid]) / 2;
  }

  function std(arr) {
    const xs = arr.filter(x => Number.isFinite(x));
    if (xs.length < 2) return null;
    const m = mean(xs);
    const v = xs.reduce((acc, x) => acc + (x - m) ** 2, 0) / (xs.length - 1);
    return Math.sqrt(v);
  }

  function cohenD(a, b) {
    const xs = a.filter(x => Number.isFinite(x));
    const ys = b.filter(x => Number.isFinite(x));
    if (xs.length < 2 || ys.length < 2) return null;
    const mx = mean(xs);
    const my = mean(ys);
    const sx = std(xs);
    const sy = std(ys);
    if (!Number.isFinite(sx) || !Number.isFinite(sy)) return null;
    const sp = Math.sqrt(((xs.length - 1) * sx * sx + (ys.length - 1) * sy * sy) / (xs.length + ys.length - 2));
    if (!Number.isFinite(sp) || sp === 0) return null;
    return (mx - my) / sp;
  }

  function pearson(x, y) {
    const pairs = [];
    for (let i = 0; i < x.length; i++) {
      const a = x[i];
      const b = y[i];
      if (Number.isFinite(a) && Number.isFinite(b)) pairs.push([a, b]);
    }
    if (pairs.length < 3) return null;
    const xs = pairs.map(p => p[0]);
    const ys = pairs.map(p => p[1]);
    const mx = mean(xs);
    const my = mean(ys);
    const sx = std(xs);
    const sy = std(ys);
    if (!sx || !sy) return null;
    let cov = 0;
    for (let i = 0; i < pairs.length; i++) cov += (pairs[i][0] - mx) * (pairs[i][1] - my);
    cov = cov / (pairs.length - 1);
    return cov / (sx * sy);
  }

  function groupCount(arr, keyFn) {
    const m = new Map();
    for (const item of arr) {
      const k = keyFn(item);
      m.set(k, (m.get(k) ?? 0) + 1);
    }
    return m;
  }

  function uniqueCount(arr, keyFn) {
    const s = new Set();
    for (const item of arr) s.add(keyFn(item));
    return s.size;
  }

  function downloadText(filename, text, mime = 'text/plain;charset=utf-8') {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function toCSV(rows, columns) {
    const cols = columns ?? (rows.length ? Object.keys(rows[0]) : []);
    const header = cols.map(c => `"${String(c).replaceAll('"', '""')}"`).join(',');
    const lines = [header];
    for (const r of rows) {
      const line = cols.map(c => {
        const v = r[c];
        if (v === null || v === undefined) return '';
        if (typeof v === 'object') return `"${JSON.stringify(v).replaceAll('"', '""')}"`;
        return `"${String(v).replaceAll('"', '""')}"`;
      }).join(',');
      lines.push(line);
    }
    return lines.join('\n');
  }

  // -----------------------------
  // Excel parsing (only 2 sheets)
  // -----------------------------
  function sheetToRows(workbook, sheetName) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { defval: null });
  }

  function detectColumns(rows) {
    const keys = new Set();
    for (const r of rows) Object.keys(r).forEach(k => keys.add(k));
    return Array.from(keys);
  }

  function makeProfileTable(rows, hideUnnamed) {
    const keys = detectColumns(rows).filter(k => {
      if (!hideUnnamed) return true;
      return !/^Unnamed/i.test(k);
    });

    const lines = [];
    lines.push('<div class="table-responsive"><table class="table table-sm table-bordered align-middle">');
    lines.push('<thead><tr><th>列</th><th>欠損</th><th>ユニーク</th><th>型（推定）</th><th>例（上位3）</th></tr></thead><tbody>');

    for (const k of keys.sort()) {
      const vals = rows.map(r => r[k]).filter(v => v !== null && v !== undefined && v !== '');
      const missing = rows.length - vals.length;
      const uniq = new Set(vals.map(v => (typeof v === 'object' ? JSON.stringify(v) : String(v)))).size;

      let typ = 'mixed';
      const numCount = vals.filter(v => Number.isFinite(Number(v))).length;
      const dateCount = vals.filter(v => {
        const d = parseISODate(v);
        return d !== null;
      }).length;
      if (vals.length === 0) typ = 'empty';
      else if (numCount / vals.length > 0.9) typ = 'number';
      else if (dateCount / vals.length > 0.9) typ = 'date';
      else typ = 'text';

      const freq = groupCount(vals, v => (typeof v === 'object' ? JSON.stringify(v) : String(v)));
      const top = Array.from(freq.entries()).sort((a, b) => b[1] - a[1]).slice(0, 3)
        .map(([val, c]) => `${escapeHtml(val)} (${c})`).join('<br/>');

      lines.push(`<tr>
        <td class="text-nowrap">${escapeHtml(k)}</td>
        <td>${missing}</td>
        <td>${uniq}</td>
        <td>${typ}</td>
        <td class="small">${top || '—'}</td>
      </tr>`);
    }

    lines.push('</tbody></table></div>');
    return lines.join('\n');
  }

  // -----------------------------
  // Normalization
  // -----------------------------
  function normalizeDiagnosis(rows) {
    // expected columns: createdAt, email, age, gender, type, axisA-D, answers_json/raw_json, interested
    return rows.map((r, idx) => {
      const createdAtRaw = r.createdAt ?? r.created_at ?? r.timestamp ?? null;
      const createdAt = parseISODate(createdAtRaw);
      const createdDate = createdAt ? toISODateString(createdAt) : null;

      const emailRaw = r.email ?? r.userEmail ?? r.mail ?? null;
      const email = (emailRaw ?? '').toString().trim();
      const emailLower = normalizeEmail(email);

      const gender = normalizeGender(r.gender);
      const type = normalizeType(r.type);

      const ageRaw = r.age;
      const ageNum = parseAgeToNumber(ageRaw);

      const axisA = safeNumber(r.axisA);
      const axisB = safeNumber(r.axisB);
      const axisC = safeNumber(r.axisC);
      const axisD = safeNumber(r.axisD);

      const interestedRaw = r.interested;
      const interested = (interestedRaw === 1 || interestedRaw === '1' || interestedRaw === true);

      // answers
      let answers = null;
      const a1 = parseJsonSafe(r.answers_json);
      if (a1 && typeof a1 === 'object' && Object.keys(a1).length) answers = a1;

      if (!answers) {
        const raw = parseJsonSafe(r.raw_json);
        if (raw && typeof raw === 'object' && raw.answers && typeof raw.answers === 'object') answers = raw.answers;
      }

      // flatten answers to numeric map (A1..D10)
      const ansFlat = {};
      if (answers && typeof answers === 'object') {
        for (const [k, v] of Object.entries(answers)) {
          const n = safeNumber(v);
          if (n !== null) ansFlat[k] = n;
        }
      }

      return {
        _row: idx,
        createdAtRaw,
        createdAt,
        createdDate,
        email,
        emailLower,
        name: r.name ?? null,
        gender,
        ageRaw,
        ageNum,
        type,
        axisA, axisB, axisC, axisD,
        interested,
        answers: ansFlat,
        raw: r
      };
    });
  }

  function normalizeReferralEvents(rows) {
    return rows.map((r, idx) => {
      const timestampRaw = r.timestamp ?? r.createdAt ?? r.time ?? null;
      const ts = parseISODate(timestampRaw);
      const date = ts ? toISODateString(ts) : null;

      const eventType = (r.eventType ?? r.type ?? r.event ?? '').toString().trim();
      const userId = (r.userId ?? '').toString().trim() || null;
      const referrerId = (r.referrerId ?? '').toString().trim() || null;

      const payload = parseJsonSafe(r.payload_json) || parseJsonSafe(r.payload) || parseJsonSafe(r.data) || null;

      const platform = payload?.platform ?? null;
      const userEmail = payload?.userEmail ?? payload?.email ?? null;
      const userEmailLower = normalizeEmail(userEmail);
      const userName = payload?.userName ?? null;
      const userType = payload?.userType ?? null;
      const gender = normalizeGender(payload?.gender);

      return {
        _row: idx,
        timestampRaw,
        ts,
        date,
        eventType,
        userId,
        referrerId,
        edge: r.edge ?? null,
        platform,
        payload,
        userEmail,
        userEmailLower,
        userName,
        userType,
        gender,
        raw: r
      };
    });
  }

  // -----------------------------
  // Derived computations
  // -----------------------------
  function buildDiagnosisUsers(diagRecords) {
    // user = emailLower grouping; for missing email, fallback to row-id based pseudo user
    const byEmail = new Map();
    for (const rec of diagRecords) {
      const key = rec.emailLower ?? `__noemail__${rec._row}`;
      if (!byEmail.has(key)) byEmail.set(key, []);
      byEmail.get(key).push(rec);
    }

    const users = [];
    for (const [key, recs] of byEmail.entries()) {
      recs.sort((a, b) => (a.createdAt?.getTime() ?? 0) - (b.createdAt?.getTime() ?? 0));
      const latest = recs[recs.length - 1];
      const favoriteRecs = recs.filter(r => r.interested);
      const hasFavorite = favoriteRecs.length > 0;
      const latestFav = hasFavorite ? favoriteRecs[favoriteRecs.length - 1] : null;

      users.push({
        userKey: key,
        emailLower: latest.emailLower,
        email: latest.email,
        latestRecord: latest,
        hasFavorite,
        favoriteCount: favoriteRecs.length,
        latestFavoriteRecord: latestFav,
        records: recs
      });
    }

    return users;
  }

  function buildReferralDerived(refEvents) {
    // daily aggregation
    const dailyMap = new Map(); // date -> {date, share, visit, complete, platformCounts}
    for (const ev of refEvents) {
      if (!ev.date) continue;
      if (!dailyMap.has(ev.date)) {
        dailyMap.set(ev.date, {
          date: ev.date,
          share: 0,
          referral_visit: 0,
          referral_complete: 0,
          platforms: new Map()
        });
      }
      const d = dailyMap.get(ev.date);
      if (ev.eventType in d) d[ev.eventType] += 1;
      if (ev.eventType === 'share') {
        const p = (ev.platform ?? 'unknown').toString();
        d.platforms.set(p, (d.platforms.get(p) ?? 0) + 1);
      }
    }

    const refDaily = Array.from(dailyMap.values()).sort((a, b) => a.date.localeCompare(b.date))
      .map(d => ({
        date: d.date,
        share: d.share,
        referral_visit: d.referral_visit,
        referral_complete: d.referral_complete,
        platforms: Object.fromEntries(d.platforms.entries())
      }));

    // referrer meta from share payloads
    const referrerMeta = new Map();
    for (const ev of refEvents) {
      if (ev.eventType !== 'share') continue;
      const rid = ev.userId || ev.referrerId;
      if (!rid) continue;
      // pick the most recent meta
      const prev = referrerMeta.get(rid);
      if (!prev || ((ev.ts?.getTime() ?? 0) > (prev._ts ?? 0))) {
        referrerMeta.set(rid, {
          referrerId: rid,
          userName: ev.userName,
          userEmail: ev.userEmail,
          userType: ev.userType,
          gender: ev.gender,
          _ts: ev.ts?.getTime() ?? 0,
          platforms: new Map()
        });
      }
      const meta = referrerMeta.get(rid);
      const p = (ev.platform ?? 'unknown').toString();
      meta.platforms.set(p, (meta.platforms.get(p) ?? 0) + 1);
    }

    // user meta from completes
    const userMeta = new Map();
    for (const ev of refEvents) {
      if (ev.eventType !== 'referral_complete') continue;
      if (!ev.userId) continue;
      const prev = userMeta.get(ev.userId);
      if (!prev || ((ev.ts?.getTime() ?? 0) > (prev._ts ?? 0))) {
        userMeta.set(ev.userId, {
          userId: ev.userId,
          userName: ev.userName,
          userEmail: ev.userEmail,
          userEmailLower: ev.userEmailLower,
          userType: ev.userType,
          gender: ev.gender,
          _ts: ev.ts?.getTime() ?? 0
        });
      }
    }

    // build edges and journeys
    const journey = new Map(); // key `${referrerId}__${userId}` -> info
    function touch(key, init) {
      if (!journey.has(key)) journey.set(key, init());
      return journey.get(key);
    }

    for (const ev of refEvents) {
      if (!ev.referrerId || !ev.userId) continue;
      const key = `${ev.referrerId}__${ev.userId}`;
      const j = touch(key, () => ({
        referrerId: ev.referrerId,
        userId: ev.userId,
        visitCount: 0,
        completeCount: 0,
        firstVisitTs: null,
        firstCompleteTs: null,
        lastTs: null
      }));

      const t = ev.ts?.getTime() ?? null;
      if (t !== null) j.lastTs = j.lastTs === null ? t : Math.max(j.lastTs, t);

      if (ev.eventType === 'referral_visit') {
        j.visitCount += 1;
        if (t !== null) j.firstVisitTs = j.firstVisitTs === null ? t : Math.min(j.firstVisitTs, t);
      } else if (ev.eventType === 'referral_complete') {
        j.completeCount += 1;
        if (t !== null) j.firstCompleteTs = j.firstCompleteTs === null ? t : Math.min(j.firstCompleteTs, t);
      }
    }

    const edgesVisitsMap = new Map();
    const edgesCompletesMap = new Map();

    for (const j of journey.values()) {
      if (j.visitCount > 0) {
        const k = `${j.referrerId}__${j.userId}`;
        edgesVisitsMap.set(k, (edgesVisitsMap.get(k) ?? 0) + j.visitCount);
      }
      if (j.completeCount > 0) {
        const k = `${j.referrerId}__${j.userId}`;
        edgesCompletesMap.set(k, (edgesCompletesMap.get(k) ?? 0) + j.completeCount);
      }
    }

    const edgesVisits = Array.from(edgesVisitsMap.entries()).map(([k, v]) => {
      const [referrerId, userId] = k.split('__');
      return { referrerId, userId, value: v };
    }).sort((a, b) => b.value - a.value);

    const edgesCompletes = Array.from(edgesCompletesMap.entries()).map(([k, v]) => {
      const [referrerId, userId] = k.split('__');
      return { referrerId, userId, value: v };
    }).sort((a, b) => b.value - a.value);

    return { refDaily, referrerMeta, userMeta, edgesVisits, edgesCompletes, journey };
  }

  function buildCompleteEmailMap(refEvents) {
    // emailLower -> latest complete {referrerId, ts}
    const m = new Map();
    for (const ev of refEvents) {
      if (ev.eventType !== 'referral_complete') continue;
      if (!ev.userEmailLower) continue;
      const t = ev.ts?.getTime() ?? 0;
      const prev = m.get(ev.userEmailLower);
      if (!prev || t > prev.ts) m.set(ev.userEmailLower, { referrerId: ev.referrerId, ts: t });
    }
    return m;
  }

  function enrichDiagnosisWithReferral(diagRecords, completeEmailToReferrer) {
    for (const rec of diagRecords) {
      const key = rec.emailLower;
      if (!key) {
        rec.referred = false;
        rec.referrerId = null;
        continue;
      }
      const info = completeEmailToReferrer.get(key);
      rec.referred = !!info;
      rec.referrerId = info?.referrerId ?? null;
      rec.referralCompleteTs = info?.ts ?? null;
    }
  }

  function buildReferrerStats(refEvents, referrerMeta, diagUserIndex) {
    // counts per referrerId based on events
    const byRef = new Map();
    function get(referrerId) {
      if (!byRef.has(referrerId)) {
        byRef.set(referrerId, {
          referrerId,
          shares: 0,
          sharesByPlatform: new Map(),
          visitUsers: new Set(),
          completeUsers: new Set(),
          completeEmails: new Set(),
          ttcHours: [],
          matchedCompletes: 0,
          matchedFavoriteUsers: 0,
          matchedFavoriteRate: null
        });
      }
      return byRef.get(referrerId);
    }

    // first visit ts per (referrer,user)
    const firstVisitTs = new Map(); // key -> ts
    for (const ev of refEvents) {
      if (!ev.referrerId || !ev.userId) continue;
      if (ev.eventType !== 'referral_visit') continue;
      const key = `${ev.referrerId}__${ev.userId}`;
      const t = ev.ts?.getTime() ?? null;
      if (t === null) continue;
      const prev = firstVisitTs.get(key);
      if (prev === undefined || t < prev) firstVisitTs.set(key, t);
    }

    for (const ev of refEvents) {
      if (ev.eventType === 'share') {
        const rid = ev.userId || ev.referrerId;
        if (!rid) continue;
        const s = get(rid);
        s.shares += 1;
        const p = (ev.platform ?? 'unknown').toString();
        s.sharesByPlatform.set(p, (s.sharesByPlatform.get(p) ?? 0) + 1);
        continue;
      }

      if (!ev.referrerId) continue;
      const s = get(ev.referrerId);

      if (ev.eventType === 'referral_visit' && ev.userId) {
        s.visitUsers.add(ev.userId);
      }

      if (ev.eventType === 'referral_complete' && ev.userId) {
        s.completeUsers.add(ev.userId);
        if (ev.userEmailLower) s.completeEmails.add(ev.userEmailLower);

        // time to complete
        const key = `${ev.referrerId}__${ev.userId}`;
        const vts = firstVisitTs.get(key);
        const cts = ev.ts?.getTime() ?? null;
        if (vts !== undefined && cts !== null && cts >= vts) {
          const hours = (cts - vts) / (1000 * 60 * 60);
          if (Number.isFinite(hours)) s.ttcHours.push(hours);
        }

        // diagnosis match / favorite
        if (ev.userEmailLower && diagUserIndex.has(ev.userEmailLower)) {
          s.matchedCompletes += 1;
          const u = diagUserIndex.get(ev.userEmailLower);
          if (u.hasFavorite) s.matchedFavoriteUsers += 1;
        }
      }
    }

    // finalize
    const rows = [];
    for (const s of byRef.values()) {
      const meta = referrerMeta.get(s.referrerId);
      const shareToVisit = s.shares > 0 ? s.visitUsers.size / s.shares : null;
      const visitToComplete = s.visitUsers.size > 0 ? s.completeUsers.size / s.visitUsers.size : null;
      const shareToComplete = s.shares > 0 ? s.completeUsers.size / s.shares : null;
      const avgTTC = mean(s.ttcHours);
      const medTTC = median(s.ttcHours);

      const matchedFavRate = s.matchedCompletes > 0 ? s.matchedFavoriteUsers / s.matchedCompletes : null;

      rows.push({
        referrerId: s.referrerId,
        referrerLabel: meta ? (meta.userName || meta.userEmail || s.referrerId) : s.referrerId,
        shares: s.shares,
        uniqueVisitors: s.visitUsers.size,
        uniqueCompletes: s.completeUsers.size,
        shareToVisit,
        visitToComplete,
        shareToComplete,
        avgTTC,
        medTTC,
        // match to diagnosis (only for completes with email)
        matchedCompletes: s.matchedCompletes,
        matchedFavoriteUsers: s.matchedFavoriteUsers,
        matchedFavRate,
        sharesByPlatform: Object.fromEntries(s.sharesByPlatform.entries())
      });
    }

    rows.sort((a, b) => (b.uniqueCompletes - a.uniqueCompletes) || (b.uniqueVisitors - a.uniqueVisitors) || (b.shares - a.shares));
    return rows;
  }

  // -----------------------------
  // Filtering
  // -----------------------------
  function withinDate(recDate, fromStr, toStr) {
    if (!recDate) return false;
    if (fromStr && recDate < fromStr) return false;
    if (toStr && recDate > toStr) return false;
    return true;
  }

  function filterDiagnosisRecords(records, filter, unit) {
    let base = records;

    if (unit === 'user') {
      // build user-level representative records (latest record for each emailLower)
      const byEmail = new Map();
      for (const r of records) {
        const k = r.emailLower ?? `__noemail__${r._row}`;
        if (!byEmail.has(k)) byEmail.set(k, []);
        byEmail.get(k).push(r);
      }
      const reps = [];
      for (const recs of byEmail.values()) {
        recs.sort((a, b) => (a.createdAt?.getTime() ?? 0) - (b.createdAt?.getTime() ?? 0));
        reps.push(recs[recs.length - 1]);
      }
      base = reps;
    }

    const from = filter.dateFrom || null;
    const to = filter.dateTo || null;
    const gender = filter.gender || 'all';
    const type = filter.type || 'all';
    const ageMin = filter.ageMin;
    const ageMax = filter.ageMax;
    const referral = filter.referral || 'all';

    return base.filter(r => {
      if (from || to) {
        const d = r.createdDate;
        if (!withinDate(d, from, to)) return false;
      }

      if (gender !== 'all') {
        if (gender === 'unknown') {
          if (r.gender === 'female' || r.gender === 'male') return false;
        } else if (r.gender !== gender) {
          return false;
        }
      }

      if (type !== 'all' && normalizeType(r.type) !== type) return false;

      if (ageMin !== null && ageMin !== undefined && ageMin !== '') {
        if (!Number.isFinite(r.ageNum) || r.ageNum < Number(ageMin)) return false;
      }
      if (ageMax !== null && ageMax !== undefined && ageMax !== '') {
        if (!Number.isFinite(r.ageNum) || r.ageNum > Number(ageMax)) return false;
      }

      if (referral !== 'all') {
        if (referral === 'referred' && !r.referred) return false;
        if (referral === 'not' && r.referred) return false;
      }

      return true;
    });
  }

  function filterFavorites(diagRecords, diagUsers, filter, unit) {
    // unit: user => diagUsers latestFavoriteRecord
    let base = [];
    if (unit === 'user') {
      base = diagUsers
        .filter(u => u.hasFavorite && u.latestFavoriteRecord)
        .map(u => u.latestFavoriteRecord);
    } else {
      base = diagRecords.filter(r => r.interested);
    }

    const from = filter.dateFrom || null;
    const to = filter.dateTo || null;
    const gender = filter.gender || 'all';
    const type = filter.type || 'all';
    const ageMin = filter.ageMin;
    const ageMax = filter.ageMax;
    const referrer = filter.referrer || 'all';

    return base.filter(r => {
      if (from || to) {
        const d = r.createdDate;
        if (!withinDate(d, from, to)) return false;
      }

      if (gender !== 'all') {
        if (gender === 'unknown') {
          if (r.gender === 'female' || r.gender === 'male') return false;
        } else if (r.gender !== gender) {
          return false;
        }
      }

      if (type !== 'all' && normalizeType(r.type) !== type) return false;

      if (ageMin !== null && ageMin !== undefined && ageMin !== '') {
        if (!Number.isFinite(r.ageNum) || r.ageNum < Number(ageMin)) return false;
      }
      if (ageMax !== null && ageMax !== undefined && ageMax !== '') {
        if (!Number.isFinite(r.ageNum) || r.ageNum > Number(ageMax)) return false;
      }

      if (referrer !== 'all') {
        if (!r.referrerId || r.referrerId !== referrer) return false;
      }

      return true;
    });
  }

  function filterReferralEvents(refEvents, filter) {
    const from = filter.dateFrom || null;
    const to = filter.dateTo || null;
    const eventType = filter.eventType || 'all';
    const platform = filter.platform || 'all';
    const referrer = filter.referrer || 'all';

    return refEvents.filter(ev => {
      if (from || to) {
        if (!withinDate(ev.date, from, to)) return false;
      }
      if (eventType !== 'all' && ev.eventType !== eventType) return false;
      if (platform !== 'all') {
        const p = (ev.platform ?? 'unknown').toString();
        if (p !== platform) return false;
      }
      if (referrer !== 'all') {
        const rid = ev.referrerId || (ev.eventType === 'share' ? (ev.userId || ev.referrerId) : null);
        if (rid !== referrer) return false;
      }
      return true;
    });
  }

  // -----------------------------
  // Rendering: helpers
  // -----------------------------
  function renderKpiCards(containerId, cards) {
    const container = document.querySelector(containerId);
    if (!container) return;
    container.innerHTML = cards.map(c => `
      <div class="col-6 col-lg-3">
        <div class="card kpi-card h-100">
          <div class="card-body">
            <div class="text-secondary small">${escapeHtml(c.label)}</div>
            <div class="fs-5 fw-semibold">${escapeHtml(c.value)}</div>
            ${c.sub ? `<div class="small text-secondary mt-1">${escapeHtml(c.sub)}</div>` : ''}
          </div>
        </div>
      </div>
    `).join('\n');
  }

  function plotEmpty(divId, message) {
    const node = el(divId);
    if (!node) return;
    Plotly.react(node, [], {
      xaxis: { visible: false },
      yaxis: { visible: false },
      annotations: [{
        text: message ?? 'データがありません',
        x: 0.5, y: 0.5, xref: 'paper', yref: 'paper',
        showarrow: false
      }],
      margin: { l: 30, r: 20, t: 20, b: 30 },
      height: node.classList.contains('plot-tall') ? 360 : 280
    }, { displayModeBar: false, responsive: true });
  }

  function plotBar(divId, x, y, title, yTitle, horiz = false) {
    const node = el(divId);
    if (!node) return;
    const data = [{
      type: 'bar',
      x: horiz ? y : x,
      y: horiz ? x : y,
      orientation: horiz ? 'h' : 'v',
      hovertemplate: horiz ? '%{y}<br>%{x}<extra></extra>' : '%{x}<br>%{y}<extra></extra>'
    }];
    const layout = {
      title: { text: title ?? '', font: { size: 12 } },
      margin: { l: horiz ? 90 : 50, r: 20, t: 30, b: 60 },
      height: node.classList.contains('plot-tall') ? 360 : 280,
      xaxis: { automargin: true, title: { text: horiz ? (yTitle ?? '') : '' } },
      yaxis: { automargin: true, title: { text: horiz ? '' : (yTitle ?? '') } }
    };
    Plotly.react(node, data, layout, { displayModeBar: false, responsive: true });
  }

  function plotLineMulti(divId, series, title) {
    const node = el(divId);
    if (!node) return;
    const data = series.map(s => ({
      type: 'scatter',
      mode: 'lines+markers',
      name: s.name,
      x: s.x,
      y: s.y,
      hovertemplate: '%{x}<br>%{y}<extra></extra>'
    }));
    const layout = {
      title: { text: title ?? '', font: { size: 12 } },
      margin: { l: 50, r: 20, t: 30, b: 50 },
      height: node.classList.contains('plot-tall') ? 360 : 280,
      xaxis: { automargin: true },
      yaxis: { automargin: true }
    };
    Plotly.react(node, data, layout, { displayModeBar: false, responsive: true });
  }

  function plotHistogram(divId, values, title, xTitle) {
    const node = el(divId);
    if (!node) return;
    const xs = values.filter(v => Number.isFinite(v));
    if (!xs.length) return plotEmpty(divId, '数値がありません');
    const data = [{
      type: 'histogram',
      x: xs,
      nbinsx: 20,
      hovertemplate: '%{x}<br>count=%{y}<extra></extra>'
    }];
    const layout = {
      title: { text: title ?? '', font: { size: 12 } },
      margin: { l: 50, r: 20, t: 30, b: 50 },
      height: node.classList.contains('plot-tall') ? 360 : 280,
      xaxis: { title: { text: xTitle ?? '' }, automargin: true },
      yaxis: { automargin: true }
    };
    Plotly.react(node, data, layout, { displayModeBar: false, responsive: true });
  }

  function plotBoxByGroup(divId, rows, valueKey, groupKey, topN = 10) {
    const node = el(divId);
    if (!node) return;
    const groups = new Map();
    for (const r of rows) {
      const g = (r[groupKey] ?? '(unknown)').toString();
      const v = r[valueKey];
      if (!Number.isFinite(v)) continue;
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(v);
    }
    const counts = Array.from(groups.entries()).map(([g, vs]) => ({ g, n: vs.length, vs }));
    counts.sort((a, b) => b.n - a.n);
    const picked = counts.slice(0, topN);

    if (!picked.length) return plotEmpty(divId, '数値がありません');

    const data = picked.map(o => ({
      type: 'box',
      name: o.g,
      y: o.vs,
      boxpoints: false
    }));

    const layout = {
      title: { text: `${valueKey}（${groupKey}別）`, font: { size: 12 } },
      margin: { l: 50, r: 20, t: 30, b: 80 },
      height: node.classList.contains('plot-tall') ? 360 : 300,
      xaxis: { tickangle: -30, automargin: true },
      yaxis: { automargin: true }
    };

    Plotly.react(node, data, layout, { displayModeBar: false, responsive: true });
  }

  function plotHeatmap(divId, labels, matrix, title) {
    const node = el(divId);
    if (!node) return;
    const data = [{
      type: 'heatmap',
      z: matrix,
      x: labels,
      y: labels,
      hovertemplate: '%{y} × %{x}<br>%{z:.3f}<extra></extra>'
    }];
    const layout = {
      title: { text: title ?? '', font: { size: 12 } },
      margin: { l: 90, r: 20, t: 30, b: 90 },
      height: node.classList.contains('plot-tall') ? 420 : 320,
    };
    Plotly.react(node, data, layout, { displayModeBar: false, responsive: true });
  }

  function plotFunnel(divId, stages, values) {
    const node = el(divId);
    if (!node) return;
    const data = [{
      type: 'funnel',
      y: stages,
      x: values,
      textinfo: 'value+percent previous'
    }];
    const layout = {
      margin: { l: 20, r: 20, t: 20, b: 20 },
      height: 280
    };
    Plotly.react(node, data, layout, { displayModeBar: false, responsive: true });
  }

  function plotSankey(divId, edges, labelFnSource, labelFnTarget, minEdge = 1) {
    const node = el(divId);
    if (!node) return;

    const filtered = edges.filter(e => e.value >= minEdge);
    if (!filtered.length) return plotEmpty(divId, 'エッジがありません（min edge を下げてください）');

    const nodes = new Map(); // label -> index
    function idx(label) {
      if (!nodes.has(label)) nodes.set(label, nodes.size);
      return nodes.get(label);
    }

    const src = [];
    const tgt = [];
    const val = [];
    for (const e of filtered) {
      const s = labelFnSource(e.referrerId);
      const t = labelFnTarget(e.userId);
      src.push(idx(s));
      tgt.push(idx(t));
      val.push(e.value);
    }

    const labels = Array.from(nodes.keys());
    const data = [{
      type: 'sankey',
      arrangement: 'snap',
      node: { label: labels, pad: 10, thickness: 14 },
      link: { source: src, target: tgt, value: val }
    }];

    const layout = {
      margin: { l: 10, r: 10, t: 10, b: 10 },
      height: node.classList.contains('plot-tall') ? 420 : 320
    };

    Plotly.react(node, data, layout, { displayModeBar: false, responsive: true });
  }

  // -----------------------------
  // Rendering: DataTables
  // -----------------------------
  function destroyDT(instance) {
    if (!instance) return null;
    try { instance.destroy(); } catch (e) { /* ignore */ }
    return null;
  }

  function renderDataTable(tableId, rows, columns, rowClick) {
    const table = document.getElementById(tableId);
    if (!table) return;

    // reset
    table.innerHTML = '';
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    for (const c of columns) {
      const th = document.createElement('th');
      th.textContent = c.title;
      tr.appendChild(th);
    }
    thead.appendChild(tr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    for (const r of rows) {
      const trb = document.createElement('tr');
      for (const c of columns) {
        const td = document.createElement('td');
        td.innerHTML = c.render ? c.render(r) : escapeHtml(r[c.key]);
        trb.appendChild(td);
      }
      if (rowClick) {
        trb.style.cursor = 'pointer';
        trb.addEventListener('click', () => rowClick(r));
      }
      tbody.appendChild(trb);
    }
    table.appendChild(tbody);

    // init DataTable
    const dt = $(table).DataTable({
      paging: true,
      pageLength: 25,
      lengthMenu: [10, 25, 50, 100],
      searching: true,
      info: true,
      order: [],
      scrollX: true
    });

    return dt;
  }

  // -----------------------------
  // Rendering: Dashboard
  // -----------------------------
  function renderDashboard() {
    const diag = state.norm.diagnosis;
    const users = state.derived.diagnosisUsers;
    const refEvents = state.derived.refEvents;

    // KPI
    const diagCount = diag.length;
    const favRecordCount = diag.filter(r => r.interested).length;
    const favUserCount = users.filter(u => u.hasFavorite).length;
    const favRate = diagCount > 0 ? favRecordCount / diagCount : null;

    const shareCount = refEvents.filter(e => e.eventType === 'share').length;
    const visitCount = refEvents.filter(e => e.eventType === 'referral_visit').length;
    const completeCount = refEvents.filter(e => e.eventType === 'referral_complete').length;

    const matchedCompletes = refEvents.filter(e => e.eventType === 'referral_complete' && e.userEmailLower && state.derived.diagnosisUserIndex.has(e.userEmailLower)).length;

    renderKpiCards('#dashKpis', [
      { label: '診断レコード', value: String(diagCount) },
      { label: 'お気に入り（レコード）', value: String(favRecordCount), sub: formatPct(favRate) },
      { label: 'お気に入りユーザー', value: String(favUserCount) },
      { label: '紹介イベント', value: String(refEvents.length), sub: `share ${shareCount} / visit ${visitCount} / complete ${completeCount}` },
    ]);

    // Trend: diagnosis + favorite by day
    const byDay = new Map();
    for (const r of diag) {
      if (!r.createdDate) continue;
      if (!byDay.has(r.createdDate)) byDay.set(r.createdDate, { date: r.createdDate, diag: 0, fav: 0 });
      const d = byDay.get(r.createdDate);
      d.diag += 1;
      if (r.interested) d.fav += 1;
    }
    const days = Array.from(byDay.values()).sort((a, b) => a.date.localeCompare(b.date));
    const x = days.map(d => d.date);
    const y1 = days.map(d => d.diag);
    const y2 = days.map(d => d.fav);
    if (x.length) {
      plotLineMulti('chartDashDiagnosisTrend', [
        { name: 'diagnosis', x, y: y1 },
        { name: 'favorites', x, y: y2 },
      ], '');
    } else {
      plotEmpty('chartDashDiagnosisTrend', '診断データがありません');
    }

    // Trend: referral events by day
    const rd = state.derived.refDaily;
    if (rd.length) {
      plotLineMulti('chartDashReferralTrend', [
        { name: 'share', x: rd.map(d => d.date), y: rd.map(d => d.share) },
        { name: 'visit', x: rd.map(d => d.date), y: rd.map(d => d.referral_visit) },
        { name: 'complete', x: rd.map(d => d.date), y: rd.map(d => d.referral_complete) },
      ], '');
    } else {
      plotEmpty('chartDashReferralTrend', '紹介イベントがありません');
    }

    // Profiles
    el('diagProfile').innerHTML = makeProfileTable(state.raw.diagnosis, state.ui.hideUnnamed);
    el('refProfile').innerHTML = makeProfileTable(state.raw.referral_events, state.ui.hideUnnamed);
  }

  // -----------------------------
  // Rendering: Diagnosis tab
  // -----------------------------
  function getDiagFilter() {
    return {
      dateFrom: el('diagDateFrom').value || null,
      dateTo: el('diagDateTo').value || null,
      gender: el('diagGender').value || 'all',
      type: el('diagType').value || 'all',
      ageMin: el('diagAgeMin').value,
      ageMax: el('diagAgeMax').value,
      referral: el('diagReferral').value || 'all'
    };
  }

  function renderDiagnosisTab() {
    const unit = el('diagUnit').value || 'record';
    const filter = getDiagFilter();
    const rows = filterDiagnosisRecords(state.norm.diagnosis, filter, unit);

    // KPI
    const total = rows.length;
    const fav = rows.filter(r => r.interested).length;
    const rate = total > 0 ? fav / total : null;
    const uniqEmail = uniqueCount(rows.filter(r => r.emailLower), r => r.emailLower);
    const ageMed = median(rows.map(r => r.ageNum));
    const femaleRate = total > 0 ? rows.filter(r => r.gender === 'female').length / total : null;

    const referredRate = total > 0 ? rows.filter(r => r.referred).length / total : null;

    renderKpiCards('#diagKpis', [
      { label: unit === 'user' ? 'ユーザー数' : '件数', value: String(total) },
      { label: 'お気に入り', value: String(fav), sub: formatPct(rate) },
      { label: 'ユニークemail', value: String(uniqEmail) },
      { label: '女性比率', value: formatPct(femaleRate), sub: `紹介あり ${formatPct(referredRate)}` },
    ]);

    // Type dist
    const typeCounts = Array.from(groupCount(rows, r => normalizeType(r.type)).entries())
      .map(([type, c]) => ({ type, c }))
      .sort((a, b) => b.c - a.c)
      .slice(0, 20);
    if (typeCounts.length) {
      plotBar('chartDiagTypeCount', typeCounts.map(d => d.type), typeCounts.map(d => d.c), '', 'count', true);
    } else plotEmpty('chartDiagTypeCount', 'データがありません');

    // Fav rate by type: compute within the (unfiltered?) Actually use rows but favorites vs all by type
    const byType = new Map();
    for (const r of rows) {
      const t = normalizeType(r.type);
      if (!byType.has(t)) byType.set(t, { t, n: 0, f: 0 });
      const o = byType.get(t);
      o.n += 1;
      if (r.interested) o.f += 1;
    }
    const favRateByType = Array.from(byType.values())
      .filter(o => o.n >= 5) // avoid tiny
      .map(o => ({ t: o.t, rate: o.f / o.n, n: o.n }))
      .sort((a, b) => b.rate - a.rate)
      .slice(0, 20);

    if (favRateByType.length) {
      plotBar('chartDiagTypeFavRate', favRateByType.map(d => `${d.t} (n=${d.n})`), favRateByType.map(d => (d.rate * 100).toFixed(1)), '', 'fav%', true);
    } else {
      plotEmpty('chartDiagTypeFavRate', '十分なデータがありません（各Type n>=5）');
    }

    // Corr heatmap
    const labels = ['axisA', 'axisB', 'axisC', 'axisD', 'age', 'interested'];
    const vectors = {
      axisA: rows.map(r => r.axisA),
      axisB: rows.map(r => r.axisB),
      axisC: rows.map(r => r.axisC),
      axisD: rows.map(r => r.axisD),
      age: rows.map(r => r.ageNum),
      interested: rows.map(r => r.interested ? 1 : 0)
    };
    const mat = [];
    for (const a of labels) {
      const row = [];
      for (const b of labels) {
        const c = pearson(vectors[a], vectors[b]);
        row.push(c === null ? null : c);
      }
      mat.push(row);
    }
    plotHeatmap('chartDiagCorr', labels, mat, '');

    // Advanced: axis distribution
    renderDiagnosisAdvanced(rows, unit);

    // Raw table
    renderDiagnosisTable(rows, unit);
  }

  function renderDiagnosisAdvanced(rows, unit) {
    // axis chart
    const axisKey = el('diagAxisPick').value || 'axisA';
    const view = el('diagAxisView').value || 'hist';
    if (view === 'hist') {
      plotHistogram('chartDiagAxis', rows.map(r => r[axisKey]), `${axisKey} 分布`, axisKey);
    } else {
      plotBoxByGroup('chartDiagAxis', rows, axisKey, 'type', 10);
    }

    // answer diff (favorites vs non)
    const favRows = rows.filter(r => r.interested);
    const nonRows = rows.filter(r => !r.interested);

    const allKeys = new Set();
    for (const r of rows) Object.keys(r.answers || {}).forEach(k => allKeys.add(k));
    const keys = Array.from(allKeys).sort();

    if (!keys.length) {
      plotEmpty('chartDiagAnswerDiff', 'answers がありません');
      setText('diagAnswersHint', 'raw_json / answers_json に answers が無い場合、このグラフは表示できません。');
    } else {
      setText('diagAnswersHint', '');
      const diffs = [];
      for (const k of keys) {
        const fv = favRows.map(r => r.answers?.[k]).filter(v => Number.isFinite(v));
        const nv = nonRows.map(r => r.answers?.[k]).filter(v => Number.isFinite(v));
        if (fv.length < 5 || nv.length < 5) continue;
        const d = mean(fv) - mean(nv);
        diffs.push({ k, d });
      }
      diffs.sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
      const top = diffs.slice(0, 10);
      if (!top.length) {
        plotEmpty('chartDiagAnswerDiff', '有効な answers が不足しています');
      } else {
        plotBar('chartDiagAnswerDiff', top.map(o => o.k), top.map(o => o.d.toFixed(2)), '', 'mean差', true);
      }
    }

    // quality
    renderDiagQuality(rows);

    // signals table (numeric)
    renderSignalTable('diagSignalTable', rows, unit);
  }

  function renderDiagQuality(rows) {
    const total = rows.length;
    if (!total) {
      el('diagQuality').innerHTML = '<div class="text-secondary small">データがありません</div>';
      return;
    }
    const missingEmail = rows.filter(r => !r.emailLower).length;
    const missingDate = rows.filter(r => !r.createdAt).length;
    const missingAge = rows.filter(r => !Number.isFinite(r.ageNum)).length;
    const outAxis = rows.filter(r => {
      const axes = [r.axisA, r.axisB, r.axisC, r.axisD];
      return axes.some(v => Number.isFinite(v) && (v < 0 || v > 100));
    }).length;

    // duplicate email (within filtered rows)
    const counts = groupCount(rows.filter(r => r.emailLower), r => r.emailLower);
    let dupEmails = 0;
    for (const c of counts.values()) if (c >= 2) dupEmails += 1;

    el('diagQuality').innerHTML = `
      <div class="small">
        <div>createdAt 欠損: <span class="fw-semibold">${missingDate}</span> / ${total}</div>
        <div>email 欠損: <span class="fw-semibold">${missingEmail}</span> / ${total}</div>
        <div>age 数値化不可: <span class="fw-semibold">${missingAge}</span> / ${total}</div>
        <div>axis 範囲外（0〜100以外）: <span class="fw-semibold">${outAxis}</span> / ${total}</div>
        <div>email 重複（ユニークemailのうち重複あり）: <span class="fw-semibold">${dupEmails}</span></div>
      </div>
    `;
  }

  function renderSignalTable(containerId, rows, unit) {
    const container = el(containerId);
    if (!container) return;

    const fav = rows.filter(r => r.interested);
    const non = rows.filter(r => !r.interested);

    if (fav.length < 10 || non.length < 10) {
      container.innerHTML = '<div class="text-secondary small">お気に入り/非お気に入りの件数が不足しています。</div>';
      return;
    }

    // numeric features: age, axes, answers
    const featureKeys = ['ageNum', 'axisA', 'axisB', 'axisC', 'axisD'];

    // include answers keys
    const ansKeys = new Set();
    for (const r of rows) Object.keys(r.answers || {}).forEach(k => ansKeys.add(k));
    const ansList = Array.from(ansKeys).sort();
    for (const k of ansList) featureKeys.push(`ans:${k}`);

    const rowsOut = [];
    for (const k of featureKeys) {
      let fvals = [];
      let nvals = [];
      if (k.startsWith('ans:')) {
        const ak = k.slice(4);
        fvals = fav.map(r => r.answers?.[ak]).filter(v => Number.isFinite(v));
        nvals = non.map(r => r.answers?.[ak]).filter(v => Number.isFinite(v));
      } else {
        fvals = fav.map(r => r[k]).filter(v => Number.isFinite(v));
        nvals = non.map(r => r[k]).filter(v => Number.isFinite(v));
      }
      if (fvals.length < 10 || nvals.length < 10) continue;
      const mf = mean(fvals);
      const mn = mean(nvals);
      const diff = mf - mn;
      const d = cohenD(fvals, nvals);
      rowsOut.push({
        feature: k.startsWith('ans:') ? k.slice(4) : k.replace('Num',''),
        mean_fav: mf,
        mean_non: mn,
        diff,
        effect_d: d,
        n_fav: fvals.length,
        n_non: nvals.length
      });
    }

    rowsOut.sort((a, b) => Math.abs(b.effect_d ?? 0) - Math.abs(a.effect_d ?? 0));
    const top = rowsOut.slice(0, 15);

    if (!top.length) {
      container.innerHTML = '<div class="text-secondary small">有効な数値特徴が不足しています。</div>';
      return;
    }

    const html = [];
    html.push('<div class="table-responsive"><table class="table table-sm table-bordered align-middle">');
    html.push('<thead><tr><th>特徴</th><th>平均（fav）</th><th>平均（non）</th><th>差</th><th>d</th><th>n</th></tr></thead><tbody>');
    for (const r of top) {
      html.push(`<tr>
        <td class="text-nowrap">${escapeHtml(r.feature)}</td>
        <td>${r.mean_fav.toFixed(2)}</td>
        <td>${r.mean_non.toFixed(2)}</td>
        <td>${r.diff.toFixed(2)}</td>
        <td>${(r.effect_d ?? 0).toFixed(2)}</td>
        <td class="small text-secondary">${r.n_fav}/${r.n_non}</td>
      </tr>`);
    }
    html.push('</tbody></table></div>');
    container.innerHTML = html.join('\n');
  }

  function renderDiagnosisTable(rows, unit) {
    // choose columns
    const cols = [
      { key: 'createdAtRaw', title: 'createdAt' },
      { key: 'type', title: 'type' },
      { key: 'gender', title: 'gender' },
      { key: 'ageRaw', title: 'age' },
      { key: 'axisA', title: 'axisA' },
      { key: 'axisB', title: 'axisB' },
      { key: 'axisC', title: 'axisC' },
      { key: 'axisD', title: 'axisD' },
      { key: 'interested', title: 'favorite', render: r => r.interested ? '1' : '0' },
      { key: 'referred', title: 'referred', render: r => r.referred ? '1' : '0' },
      { key: 'referrerId', title: 'referrer', render: r => {
        if (!r.referrerId) return '';
        return state.ui.maskPII ? maskId('r', r.referrerId) : escapeHtml(r.referrerId);
      }},
      { key: 'email', title: 'email', render: r => {
        if (!r.email) return '';
        return state.ui.maskPII ? maskId('u', r.emailLower || r.email) : escapeHtml(r.email);
      }},
      { key: '_row', title: 'row' }
    ];

    state.dt.tableDiagnosis = destroyDT(state.dt.tableDiagnosis);
    state.dt.tableDiagnosis = renderDataTable('tableDiagnosis', rows, cols);
  }

  // -----------------------------
  // Rendering: Favorites tab
  // -----------------------------
  function getFavFilter() {
    return {
      dateFrom: el('favDateFrom').value || null,
      dateTo: el('favDateTo').value || null,
      gender: el('favGender').value || 'all',
      type: el('favType').value || 'all',
      ageMin: el('favAgeMin').value,
      ageMax: el('favAgeMax').value,
      referrer: el('favReferrer').value || 'all'
    };
  }

  function buildNonFavoriteBaseline(unit, filter) {
    // for comparison in Favorites tab:
    // unit=user => users with no favorite -> latestRecord
    // unit=record => diagnosis records with interested=false
    const diag = state.norm.diagnosis;
    if (unit === 'user') {
      const users = state.derived.diagnosisUsers.filter(u => !u.hasFavorite && u.latestRecord);
      // apply filter on latestRecord
      return users.map(u => u.latestRecord).filter(r => {
        // reuse favorites filter logic but without referrer restriction (still meaningful though)
        const f = { ...filter, referrer: 'all' };
        return filterFavorites([r], [{hasFavorite:false, latestFavoriteRecord:null}], f, 'record').length > 0;
      });
    }
    // record
    return diag.filter(r => !r.interested).filter(r => {
      // use same filter but ignoring referrer filter
      const f = { ...filter, referrer: 'all' };
      return filterFavorites([r], [], f, 'record').length > 0;
    });
  }

  function renderFavoritesTab() {
    const unit = el('favUnit').value || 'user';
    const filter = getFavFilter();

    const favRows = filterFavorites(state.norm.diagnosis, state.derived.diagnosisUsers, filter, unit);
    const nonRows = buildNonFavoriteBaseline(unit, filter);

    // KPI
    const favN = favRows.length;
    const allN = unit === 'user'
      ? state.derived.diagnosisUsers.filter(u => u.emailLower).length
      : state.norm.diagnosis.length;

    const favRate = allN > 0 ? favN / allN : null;

    const referredFav = favRows.filter(r => r.referred).length;
    const referredRate = favN > 0 ? referredFav / favN : null;

    const topType = (() => {
      const m = groupCount(favRows, r => normalizeType(r.type));
      const arr = Array.from(m.entries()).sort((a, b) => b[1] - a[1]);
      return arr.length ? `${arr[0][0]} (${arr[0][1]})` : '—';
    })();

    renderKpiCards('#favKpis', [
      { label: unit === 'user' ? 'お気に入りユーザー' : 'お気に入りレコード', value: String(favN), sub: `${formatPct(favRate)}（全体比）` },
      { label: '紹介あり（complete）', value: String(referredFav), sub: formatPct(referredRate) },
      { label: 'Top Type', value: topType },
      { label: '比較母数（非お気に入り）', value: String(nonRows.length) }
    ]);

    // Trend
    const byDay = new Map();
    for (const r of favRows) {
      if (!r.createdDate) continue;
      byDay.set(r.createdDate, (byDay.get(r.createdDate) ?? 0) + 1);
    }
    const days = Array.from(byDay.entries()).sort((a, b) => a[0].localeCompare(b[0]));
    if (days.length) {
      plotLineMulti('chartFavTrend', [{ name: 'favorites', x: days.map(d => d[0]), y: days.map(d => d[1]) }], '');
    } else plotEmpty('chartFavTrend', 'データがありません');

    // Type distribution
    const typeCounts = Array.from(groupCount(favRows, r => normalizeType(r.type)).entries())
      .map(([type, c]) => ({ type, c }))
      .sort((a, b) => b.c - a.c)
      .slice(0, 20);
    if (typeCounts.length) {
      plotBar('chartFavTypeCount', typeCounts.map(d => d.type), typeCounts.map(d => d.c), '', 'count', true);
    } else plotEmpty('chartFavTypeCount', 'データがありません');

    // Axis diff
    const axisKeys = ['axisA', 'axisB', 'axisC', 'axisD'];
    const diffs = axisKeys.map(k => {
      const fv = favRows.map(r => r[k]).filter(v => Number.isFinite(v));
      const nv = nonRows.map(r => r[k]).filter(v => Number.isFinite(v));
      const d = (mean(fv) ?? 0) - (mean(nv) ?? 0);
      return { k, d };
    });
    plotBar('chartFavAxisDiff', diffs.map(o => o.k), diffs.map(o => o.d.toFixed(2)), '', 'mean差', true);

    // Answer diff top 10
    const allKeys = new Set();
    for (const r of favRows) Object.keys(r.answers || {}).forEach(k => allKeys.add(k));
    for (const r of nonRows) Object.keys(r.answers || {}).forEach(k => allKeys.add(k));
    const keys = Array.from(allKeys).sort();

    if (!keys.length) {
      plotEmpty('chartFavAnswerDiff', 'answers がありません');
      setText('favAnswersHint', 'raw_json / answers_json に answers が無い場合、このグラフは表示できません。');
    } else {
      setText('favAnswersHint', '');
      const out = [];
      for (const k of keys) {
        const fv = favRows.map(r => r.answers?.[k]).filter(v => Number.isFinite(v));
        const nv = nonRows.map(r => r.answers?.[k]).filter(v => Number.isFinite(v));
        if (fv.length < 10 || nv.length < 10) continue;
        out.push({ k, d: mean(fv) - mean(nv) });
      }
      out.sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
      const top = out.slice(0, 10);
      if (!top.length) plotEmpty('chartFavAnswerDiff', '有効な answers が不足しています');
      else plotBar('chartFavAnswerDiff', top.map(o => o.k), top.map(o => o.d.toFixed(2)), '', 'mean差', true);
    }

    // Signal table
    renderFavSignalTable(favRows, nonRows);

    // Favorites table
    renderFavoritesTable(favRows, unit);
  }

  function renderFavSignalTable(favRows, nonRows) {
    const container = el('favSignalTable');
    if (!container) return;

    if (favRows.length < 10 || nonRows.length < 10) {
      container.innerHTML = '<div class="text-secondary small">比較に必要な件数が不足しています。</div>';
      return;
    }

    const featureKeys = ['ageNum', 'axisA', 'axisB', 'axisC', 'axisD'];

    const ansKeys = new Set();
    for (const r of favRows) Object.keys(r.answers || {}).forEach(k => ansKeys.add(k));
    for (const r of nonRows) Object.keys(r.answers || {}).forEach(k => ansKeys.add(k));
    const ansList = Array.from(ansKeys).sort();
    for (const k of ansList) featureKeys.push(`ans:${k}`);

    const rowsOut = [];
    for (const k of featureKeys) {
      let fvals = [];
      let nvals = [];
      if (k.startsWith('ans:')) {
        const ak = k.slice(4);
        fvals = favRows.map(r => r.answers?.[ak]).filter(v => Number.isFinite(v));
        nvals = nonRows.map(r => r.answers?.[ak]).filter(v => Number.isFinite(v));
      } else {
        fvals = favRows.map(r => r[k]).filter(v => Number.isFinite(v));
        nvals = nonRows.map(r => r[k]).filter(v => Number.isFinite(v));
      }
      if (fvals.length < 10 || nvals.length < 10) continue;
      const mf = mean(fvals);
      const mn = mean(nvals);
      const diff = mf - mn;
      const d = cohenD(fvals, nvals);
      rowsOut.push({
        feature: k.startsWith('ans:') ? k.slice(4) : k.replace('Num',''),
        mean_fav: mf,
        mean_non: mn,
        diff,
        effect_d: d,
        n_fav: fvals.length,
        n_non: nvals.length
      });
    }

    rowsOut.sort((a, b) => Math.abs(b.effect_d ?? 0) - Math.abs(a.effect_d ?? 0));
    const top = rowsOut.slice(0, 20);

    if (!top.length) {
      container.innerHTML = '<div class="text-secondary small">有効な数値特徴が不足しています。</div>';
      return;
    }

    const html = [];
    html.push('<div class="table-responsive"><table class="table table-sm table-bordered align-middle">');
    html.push('<thead><tr><th>特徴</th><th>平均（fav）</th><th>平均（non）</th><th>差</th><th>d</th><th>n</th></tr></thead><tbody>');
    for (const r of top) {
      html.push(`<tr>
        <td class="text-nowrap">${escapeHtml(r.feature)}</td>
        <td>${r.mean_fav.toFixed(2)}</td>
        <td>${r.mean_non.toFixed(2)}</td>
        <td>${r.diff.toFixed(2)}</td>
        <td>${(r.effect_d ?? 0).toFixed(2)}</td>
        <td class="small text-secondary">${r.n_fav}/${r.n_non}</td>
      </tr>`);
    }
    html.push('</tbody></table></div>');
    container.innerHTML = html.join('\n');
  }

  function renderFavoritesTable(rows, unit) {
    const cols = [
      { key: 'createdAtRaw', title: 'createdAt' },
      { key: 'type', title: 'type' },
      { key: 'gender', title: 'gender' },
      { key: 'ageRaw', title: 'age' },
      { key: 'axisA', title: 'axisA' },
      { key: 'axisB', title: 'axisB' },
      { key: 'axisC', title: 'axisC' },
      { key: 'axisD', title: 'axisD' },
      { key: 'referred', title: 'referred', render: r => r.referred ? '1' : '0' },
      { key: 'referrerId', title: 'referrer', render: r => {
        if (!r.referrerId) return '';
        return state.ui.maskPII ? maskId('r', r.referrerId) : escapeHtml(r.referrerId);
      }},
      { key: 'email', title: 'user', render: r => {
        const key = r.emailLower || `row-${r._row}`;
        return state.ui.maskPII ? maskId('u', key) : escapeHtml(r.email || key);
      }},
      { key: '_row', title: 'row' }
    ];

    state.dt.tableFavorites = destroyDT(state.dt.tableFavorites);
    state.dt.tableFavorites = renderDataTable('tableFavorites', rows, cols);
  }

  // -----------------------------
  // Rendering: Referral tab
  // -----------------------------
  function getRefFilter() {
    return {
      dateFrom: el('refDateFrom').value || null,
      dateTo: el('refDateTo').value || null,
      eventType: el('refEventType').value || 'all',
      platform: el('refPlatform').value || 'all',
      referrer: el('refReferrer').value || 'all'
    };
  }

  function renderReferralTab() {
    const filter = getRefFilter();
    const rows = filterReferralEvents(state.derived.refEvents, filter);

    // KPI
    const share = rows.filter(e => e.eventType === 'share').length;
    const visit = rows.filter(e => e.eventType === 'referral_visit').length;
    const complete = rows.filter(e => e.eventType === 'referral_complete').length;

    // uniques by id (visitors and completes)
    const uniqVisitors = uniqueCount(rows.filter(e => e.eventType === 'referral_visit' && e.userId), e => e.userId);
    const uniqCompletes = uniqueCount(rows.filter(e => e.eventType === 'referral_complete' && e.userId), e => e.userId);

    const sv = share > 0 ? uniqVisitors / share : null;
    const vc = uniqVisitors > 0 ? uniqCompletes / uniqVisitors : null;
    const sc = share > 0 ? uniqCompletes / share : null;

    // favorites among matched completes (email join)
    const completeEmails = rows.filter(e => e.eventType === 'referral_complete' && e.userEmailLower).map(e => e.userEmailLower);
    const uniqCompleteEmails = new Set(completeEmails);
    let matched = 0;
    let matchedFav = 0;
    for (const emailLower of uniqCompleteEmails.values()) {
      if (state.derived.diagnosisUserIndex.has(emailLower)) {
        matched += 1;
        if (state.derived.diagnosisUserIndex.get(emailLower).hasFavorite) matchedFav += 1;
      }
    }
    const matchedFavRate = matched > 0 ? matchedFav / matched : null;

    renderKpiCards('#refKpis', [
      { label: 'share', value: String(share) },
      { label: 'visit（unique）', value: String(uniqVisitors), sub: share > 0 ? `share→visit ${formatPct(sv)}` : '—' },
      { label: 'complete（unique）', value: String(uniqCompletes), sub: uniqVisitors > 0 ? `visit→complete ${formatPct(vc)}` : '—' },
      { label: 'complete→fav（診断マッチ）', value: `${matchedFav}/${matched}`, sub: matched > 0 ? formatPct(matchedFavRate) : '—' },
    ]);

    // Funnel chart (use totals from filtered rows)
    plotFunnel('chartRefFunnel', ['share', 'visit (unique)', 'complete (unique)'], [share, uniqVisitors, uniqCompletes]);

    // Trend chart: day counts within filter
    const dailyMap = new Map();
    for (const ev of rows) {
      if (!ev.date) continue;
      if (!dailyMap.has(ev.date)) dailyMap.set(ev.date, { date: ev.date, share: 0, visit: 0, complete: 0 });
      const d = dailyMap.get(ev.date);
      if (ev.eventType === 'share') d.share += 1;
      if (ev.eventType === 'referral_visit') d.visit += 1;
      if (ev.eventType === 'referral_complete') d.complete += 1;
    }
    const days = Array.from(dailyMap.values()).sort((a, b) => a.date.localeCompare(b.date));
    if (days.length) {
      plotLineMulti('chartRefTrend', [
        { name: 'share', x: days.map(d => d.date), y: days.map(d => d.share) },
        { name: 'visit', x: days.map(d => d.date), y: days.map(d => d.visit) },
        { name: 'complete', x: days.map(d => d.date), y: days.map(d => d.complete) },
      ], '');
    } else plotEmpty('chartRefTrend', 'データがありません');

    // Referrer leaderboard (computed on all events, then filtered? We'll compute on ALL but allow filter by date/platform/referrer? For simplicity: use filtered rows, but this can hide global context.)
    renderReferrerLeaderboard(filter);

    // Sankey
    renderReferralSankey();

    // Raw events table
    renderReferralEventsTable(rows);

    // Selected referrer details
    renderSelectedReferrerDetails();
  }

  function renderReferrerLeaderboard(filter) {
    // compute stats from refEvents filtered by date range only (ignore eventType/platform selection to keep stable)
    const baseFilter = { ...filter, eventType: 'all', platform: 'all', referrer: 'all' };
    const baseEvents = filterReferralEvents(state.derived.refEvents, baseFilter);

    const stats = buildReferrerStats(baseEvents, state.derived.referrerMeta, state.derived.diagnosisUserIndex);

    const cols = [
      { key: 'referrerLabel', title: 'referrer', render: r => {
        const label = state.ui.maskPII ? maskId('r', r.referrerId) : r.referrerLabel;
        return escapeHtml(label);
      }},
      { key: 'shares', title: 'shares' },
      { key: 'uniqueVisitors', title: 'visitors' },
      { key: 'uniqueCompletes', title: 'completes' },
      { key: 'visitToComplete', title: 'visit→complete', render: r => formatPct(r.visitToComplete ?? NaN) },
      { key: 'avgTTC', title: 'avg h', render: r => Number.isFinite(r.avgTTC) ? r.avgTTC.toFixed(1) : '—' },
      { key: 'matchedFavRate', title: 'fav rate', render: r => Number.isFinite(r.matchedFavRate) ? formatPct(r.matchedFavRate) : '—' },
    ];

    state.dt.tableReferrers = destroyDT(state.dt.tableReferrers);
    state.dt.tableReferrers = renderDataTable('tableReferrers', stats, cols, (row) => {
      state.selection.selectedReferrerId = row.referrerId;
      // also update dropdown selection (so filters sync)
      const sel = el('refReferrer');
      if (sel) sel.value = row.referrerId;
      renderSelectedReferrerDetails();
      // open details
      // (no forced open)
    });
  }

  function renderReferralSankey() {
    const mode = el('refSankeyMode').value || 'visits';
    const minEdge = Number(el('refMinEdge').value || '1');

    const edges = mode === 'completes' ? state.derived.edges.completes : state.derived.edges.visits;

    const labelSource = (rid) => {
      if (!rid) return 'referrer:unknown';
      if (state.ui.maskPII) return `R:${maskId('r', rid)}`;
      const meta = state.derived.referrerMeta.get(rid);
      return meta?.userEmail || meta?.userName || rid;
    };
    const labelTarget = (uid) => {
      if (!uid) return 'user:unknown';
      if (state.ui.maskPII) return `U:${maskId('u', uid)}`;
      const meta = state.derived.userMeta.get(uid);
      return meta?.userEmail || meta?.userName || uid;
    };

    plotSankey('chartRefSankey', edges, labelSource, labelTarget, minEdge);

    const shown = edges.filter(e => e.value >= minEdge).length;
    setText('refSankeyHint', `表示エッジ: ${shown}（min edge=${minEdge}）`);
  }

  function renderReferralEventsTable(rows) {
    const cols = [
      { key: 'timestampRaw', title: 'timestamp' },
      { key: 'eventType', title: 'eventType' },
      { key: 'platform', title: 'platform', render: r => escapeHtml((r.platform ?? '')) },
      { key: 'referrerId', title: 'referrer', render: r => {
        const rid = (r.eventType === 'share') ? (r.userId || r.referrerId) : r.referrerId;
        if (!rid) return '';
        return state.ui.maskPII ? maskId('r', rid) : escapeHtml(rid);
      }},
      { key: 'userId', title: 'userId', render: r => {
        if (!r.userId) return '';
        return state.ui.maskPII ? maskId('u', r.userId) : escapeHtml(r.userId);
      }},
      { key: 'userEmail', title: 'userEmail', render: r => {
        if (!r.userEmail) return '';
        return state.ui.maskPII ? maskId('u', r.userEmailLower || r.userEmail) : escapeHtml(r.userEmail);
      }},
      { key: '_row', title: 'row' }
    ];

    state.dt.tableRefEvents = destroyDT(state.dt.tableRefEvents);
    state.dt.tableRefEvents = renderDataTable('tableRefEvents', rows, cols);
  }

  function renderSelectedReferrerDetails() {
    const rid = state.selection.selectedReferrerId || el('refReferrer').value;
    if (!rid || rid === 'all') {
      el('refSelectedSummary').innerHTML = '<span class="text-secondary">紹介者を選択してください。</span>';
      plotEmpty('chartRefTTC', '紹介者を選択してください');
      state.dt.tableRefEdges = destroyDT(state.dt.tableRefEdges);
      el('tableRefEdges').innerHTML = '';
      return;
    }

    // build details from all events (respect date filter range but not eventType/platform)
    const filter = getRefFilter();
    const baseFilter = { ...filter, eventType: 'all', platform: 'all', referrer: rid };
    const evs = filterReferralEvents(state.derived.refEvents, baseFilter);

    const shares = evs.filter(e => e.eventType === 'share').length;
    const visits = evs.filter(e => e.eventType === 'referral_visit').map(e => e.userId).filter(Boolean);
    const completes = evs.filter(e => e.eventType === 'referral_complete').map(e => e.userId).filter(Boolean);
    const uniqV = new Set(visits).size;
    const uniqC = new Set(completes).size;

    // time-to-complete distribution
    const ttc = [];
    // compute per userId: first visit and first complete
    const firstVisit = new Map();
    for (const e of evs) {
      if (e.eventType !== 'referral_visit' || !e.userId) continue;
      const t = e.ts?.getTime() ?? null;
      if (t === null) continue;
      const prev = firstVisit.get(e.userId);
      if (prev === undefined || t < prev) firstVisit.set(e.userId, t);
    }
    for (const e of evs) {
      if (e.eventType !== 'referral_complete' || !e.userId) continue;
      const t = e.ts?.getTime() ?? null;
      const v = firstVisit.get(e.userId);
      if (t === null || v === undefined || t < v) continue;
      ttc.push((t - v) / (1000 * 60 * 60));
    }

    // matched favorites among complete emails (unique)
    const completeEmails = evs.filter(e => e.eventType === 'referral_complete' && e.userEmailLower).map(e => e.userEmailLower);
    const uniqCompleteEmails = new Set(completeEmails);
    let matched = 0;
    let matchedFav = 0;
    for (const emailLower of uniqCompleteEmails.values()) {
      if (state.derived.diagnosisUserIndex.has(emailLower)) {
        matched += 1;
        if (state.derived.diagnosisUserIndex.get(emailLower).hasFavorite) matchedFav += 1;
      }
    }

    const meta = state.derived.referrerMeta.get(rid);
    const label = state.ui.maskPII ? maskId('r', rid) : (meta?.userEmail || meta?.userName || rid);

    el('refSelectedSummary').innerHTML = `
      <div class="small">
        <div><span class="text-secondary">referrer:</span> <span class="fw-semibold">${escapeHtml(label)}</span></div>
        <div class="mt-2">shares: <span class="fw-semibold">${shares}</span></div>
        <div>visitors（unique userId）: <span class="fw-semibold">${uniqV}</span></div>
        <div>completes（unique userId）: <span class="fw-semibold">${uniqC}</span></div>
        <div class="mt-2">診断マッチ（complete email）: <span class="fw-semibold">${matched}</span> / fav: <span class="fw-semibold">${matchedFav}</span>（${matched ? formatPct(matchedFav / matched) : '—'}）</div>
      </div>
    `;

    if (ttc.length) plotHistogram('chartRefTTC', ttc, '', 'hours');
    else plotEmpty('chartRefTTC', 'time-to-complete がありません');

    // edges table for this referrer (visits & completes)
    const byUser = new Map();
    for (const e of evs) {
      if (!e.userId) continue;
      if (!byUser.has(e.userId)) byUser.set(e.userId, { userId: e.userId, visits: 0, completes: 0, firstVisit: null, firstComplete: null, hours: null, email: null });
      const o = byUser.get(e.userId);
      const t = e.ts?.getTime() ?? null;

      if (e.eventType === 'referral_visit') {
        o.visits += 1;
        if (t !== null) o.firstVisit = o.firstVisit === null ? t : Math.min(o.firstVisit, t);
      } else if (e.eventType === 'referral_complete') {
        o.completes += 1;
        if (t !== null) o.firstComplete = o.firstComplete === null ? t : Math.min(o.firstComplete, t);
        if (e.userEmailLower) o.email = e.userEmailLower;
      }
    }

    const edgeRows = [];
    for (const o of byUser.values()) {
      if (o.firstVisit !== null && o.firstComplete !== null && o.firstComplete >= o.firstVisit) {
        o.hours = (o.firstComplete - o.firstVisit) / (1000 * 60 * 60);
      }
      edgeRows.push(o);
    }
    edgeRows.sort((a, b) => (b.completes - a.completes) || (b.visits - a.visits));

    const cols = [
      { key: 'userId', title: 'user', render: r => state.ui.maskPII ? maskId('u', r.userId) : escapeHtml(r.userId) },
      { key: 'visits', title: 'visits' },
      { key: 'completes', title: 'completes' },
      { key: 'hours', title: 'hours', render: r => Number.isFinite(r.hours) ? r.hours.toFixed(2) : '—' },
      { key: 'email', title: 'diag match', render: r => {
        if (!r.email) return '';
        const ok = state.derived.diagnosisUserIndex.has(r.email);
        const fav = ok && state.derived.diagnosisUserIndex.get(r.email).hasFavorite;
        const label2 = state.ui.maskPII ? maskId('u', r.email) : r.email;
        return `${escapeHtml(label2)} ${ok ? (fav ? '★' : '✓') : ''}`;
      }}
    ];

    state.dt.tableRefEdges = destroyDT(state.dt.tableRefEdges);
    state.dt.tableRefEdges = renderDataTable('tableRefEdges', edgeRows, cols);
  }

  // -----------------------------
  // UI: populate options
  // -----------------------------
  function populateTypeOptions(selectId, types) {
    const sel = el(selectId);
    if (!sel) return;
    const current = sel.value || 'all';
    sel.innerHTML = '<option value="all">すべて</option>';
    for (const t of types) {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t;
      sel.appendChild(opt);
    }
    sel.value = types.includes(current) ? current : 'all';
  }

  function populatePlatformOptions(platforms) {
    const sel = el('refPlatform');
    if (!sel) return;
    const current = sel.value || 'all';
    sel.innerHTML = '<option value="all">すべて</option>';
    for (const p of platforms) {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      sel.appendChild(opt);
    }
    sel.value = platforms.includes(current) ? current : 'all';
  }

  function populateReferrerOptions() {
    const ids = Array.from(state.derived.referrerMeta.keys()).sort();
    // If meta is missing, also include from edges
    for (const e of state.derived.edges.visits) ids.push(e.referrerId);
    for (const e of state.derived.edges.completes) ids.push(e.referrerId);

    const uniq = Array.from(new Set(ids)).sort();

    const renderLabel = (rid) => {
      const meta = state.derived.referrerMeta.get(rid);
      if (state.ui.maskPII) return maskId('r', rid);
      return meta?.userEmail || meta?.userName || rid;
    };

    const selRef = el('refReferrer');
    const selFav = el('favReferrer');
    if (selRef) {
      const current = selRef.value || 'all';
      selRef.innerHTML = '<option value="all">すべて</option>';
      for (const rid of uniq) {
        const opt = document.createElement('option');
        opt.value = rid;
        opt.textContent = renderLabel(rid);
        selRef.appendChild(opt);
      }
      if (current !== 'all' && uniq.includes(current)) selRef.value = current;
      else selRef.value = 'all';
    }
    if (selFav) {
      const current = selFav.value || 'all';
      selFav.innerHTML = '<option value="all">すべて</option>';
      for (const rid of uniq) {
        const opt = document.createElement('option');
        opt.value = rid;
        opt.textContent = renderLabel(rid);
        selFav.appendChild(opt);
      }
      if (current !== 'all' && uniq.includes(current)) selFav.value = current;
      else selFav.value = 'all';
    }
  }

  function setDefaultDateFilters() {
    // diagnosis
    const diagDates = state.norm.diagnosis.map(r => r.createdDate).filter(Boolean).sort();
    const refDates = state.derived.refEvents.map(r => r.date).filter(Boolean).sort();

    if (diagDates.length) {
      el('diagDateFrom').value = diagDates[0];
      el('diagDateTo').value = diagDates[diagDates.length - 1];
      el('favDateFrom').value = diagDates[0];
      el('favDateTo').value = diagDates[diagDates.length - 1];
    }
    if (refDates.length) {
      el('refDateFrom').value = refDates[0];
      el('refDateTo').value = refDates[refDates.length - 1];
    }
  }

  // -----------------------------
  // Load workbook
  // -----------------------------
  async function loadWorkbookFromFile(file) {
    state.file = file;
    setText('fileName', file?.name ?? '—');
    setText('loadStatus', '読み込み中…');

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });

    const diagSheetName = (wb.SheetNames || []).find(n => (n || '').toString().toLowerCase() === 'diagnosis');
    const refSheetName = (wb.SheetNames || []).find(n => (n || '').toString().toLowerCase() === 'referral_events');

    if (!diagSheetName || !refSheetName) {
      setText('loadStatus', '必要なシートが見つかりません');
      alert('このファイルには diagnosis と referral_events シートが必要です。');
      return;
    }

    state.raw.diagnosis = sheetToRows(wb, diagSheetName);
    state.raw.referral_events = sheetToRows(wb, refSheetName);

    state.norm.diagnosis = normalizeDiagnosis(state.raw.diagnosis);
    state.norm.referral_events = normalizeReferralEvents(state.raw.referral_events);

    // derived
    state.derived.diagnosisUsers = buildDiagnosisUsers(state.norm.diagnosis);
    state.derived.diagnosisUserIndex = new Map();
    for (const u of state.derived.diagnosisUsers) {
      if (u.emailLower) state.derived.diagnosisUserIndex.set(u.emailLower, u);
    }

    state.derived.refEvents = state.norm.referral_events;
    const refD = buildReferralDerived(state.derived.refEvents);
    state.derived.refDaily = refD.refDaily;
    state.derived.referrerMeta = refD.referrerMeta;
    state.derived.userMeta = refD.userMeta;
    state.derived.edges.visits = refD.edgesVisits;
    state.derived.edges.completes = refD.edgesCompletes;
    state.derived.journey = refD.journey;

    state.derived.completeEmailToReferrer = buildCompleteEmailMap(state.derived.refEvents);

    // enrich diagnosis with referral mapping
    enrichDiagnosisWithReferral(state.norm.diagnosis, state.derived.completeEmailToReferrer);

    // refresh users (because latestRecord got referral fields updated)
    state.derived.diagnosisUsers = buildDiagnosisUsers(state.norm.diagnosis);
    state.derived.diagnosisUserIndex = new Map();
    for (const u of state.derived.diagnosisUsers) {
      if (u.emailLower) state.derived.diagnosisUserIndex.set(u.emailLower, u);
    }

    // UI options
    const types = Array.from(new Set(state.norm.diagnosis.map(r => normalizeType(r.type)))).sort();
    populateTypeOptions('diagType', types);
    populateTypeOptions('favType', types);

    const platforms = Array.from(new Set(state.derived.refEvents.filter(e => e.eventType === 'share').map(e => (e.platform ?? 'unknown').toString()))).sort();
    populatePlatformOptions(platforms);

    populateReferrerOptions();

    setDefaultDateFilters();

    // Counts
    setText('diagCount', String(state.norm.diagnosis.length));
    setText('favCount', String(state.norm.diagnosis.filter(r => r.interested).length));
    setText('refCount', String(state.derived.refEvents.length));

    setText('loadStatus', '完了');

    hide('noDataHint');

    // render
    renderAll();
  }

  function renderAll() {
    renderDashboard();
    renderDiagnosisTab();
    renderFavoritesTab();
    renderReferralTab();
  }

  // -----------------------------
  // Export buttons
  // -----------------------------
  function setupExportButtons() {
    el('btnExportDiagCSV').addEventListener('click', () => {
      const unit = el('diagUnit').value || 'record';
      const rows = filterDiagnosisRecords(state.norm.diagnosis, getDiagFilter(), unit);
      const out = rows.map(r => ({
        createdAt: r.createdAtRaw,
        type: r.type,
        gender: r.gender,
        age: r.ageRaw,
        axisA: r.axisA, axisB: r.axisB, axisC: r.axisC, axisD: r.axisD,
        favorite: r.interested ? 1 : 0,
        referred: r.referred ? 1 : 0,
        referrerId: r.referrerId,
        email: state.ui.maskPII ? maskId('u', r.emailLower || r.email || `row-${r._row}`) : r.email
      }));
      downloadText('diagnosis_filtered.csv', toCSV(out));
    });

    el('btnExportDiagJSON').addEventListener('click', () => {
      const unit = el('diagUnit').value || 'record';
      const rows = filterDiagnosisRecords(state.norm.diagnosis, getDiagFilter(), unit);
      const out = rows.map(r => ({
        createdAt: r.createdAtRaw,
        type: r.type,
        gender: r.gender,
        age: r.ageRaw,
        axes: { axisA: r.axisA, axisB: r.axisB, axisC: r.axisC, axisD: r.axisD },
        favorite: r.interested ? 1 : 0,
        referred: r.referred ? 1 : 0,
        referrerId: r.referrerId,
        user: state.ui.maskPII ? maskId('u', r.emailLower || r.email || `row-${r._row}`) : (r.email || null),
        answers: r.answers
      }));
      downloadText('diagnosis_filtered.json', JSON.stringify(out, null, 2), 'application/json;charset=utf-8');
    });

    el('btnExportFavCSV').addEventListener('click', () => {
      const unit = el('favUnit').value || 'user';
      const rows = filterFavorites(state.norm.diagnosis, state.derived.diagnosisUsers, getFavFilter(), unit);
      const out = rows.map(r => ({
        createdAt: r.createdAtRaw,
        type: r.type,
        gender: r.gender,
        age: r.ageRaw,
        axisA: r.axisA, axisB: r.axisB, axisC: r.axisC, axisD: r.axisD,
        referred: r.referred ? 1 : 0,
        referrerId: r.referrerId,
        user: state.ui.maskPII ? maskId('u', r.emailLower || r.email || `row-${r._row}`) : (r.email || null)
      }));
      downloadText('favorites_filtered.csv', toCSV(out));
    });

    el('btnExportFavJSON').addEventListener('click', () => {
      const unit = el('favUnit').value || 'user';
      const rows = filterFavorites(state.norm.diagnosis, state.derived.diagnosisUsers, getFavFilter(), unit);
      const out = rows.map(r => ({
        createdAt: r.createdAtRaw,
        type: r.type,
        gender: r.gender,
        age: r.ageRaw,
        axes: { axisA: r.axisA, axisB: r.axisB, axisC: r.axisC, axisD: r.axisD },
        referred: r.referred ? 1 : 0,
        referrerId: r.referrerId,
        user: state.ui.maskPII ? maskId('u', r.emailLower || r.email || `row-${r._row}`) : (r.email || null),
        answers: r.answers
      }));
      downloadText('favorites_filtered.json', JSON.stringify(out, null, 2), 'application/json;charset=utf-8');
    });

    el('btnExportRefCSV').addEventListener('click', () => {
      const rows = filterReferralEvents(state.derived.refEvents, getRefFilter());
      const out = rows.map(e => ({
        timestamp: e.timestampRaw,
        eventType: e.eventType,
        platform: e.platform,
        referrerId: e.eventType === 'share' ? (e.userId || e.referrerId) : e.referrerId,
        userId: e.userId,
        userEmail: state.ui.maskPII ? (e.userEmailLower ? maskId('u', e.userEmailLower) : '') : (e.userEmail || '')
      }));
      downloadText('referral_events_filtered.csv', toCSV(out));
    });

    el('btnExportRefJSON').addEventListener('click', () => {
      const rows = filterReferralEvents(state.derived.refEvents, getRefFilter());
      const out = rows.map(e => ({
        timestamp: e.timestampRaw,
        eventType: e.eventType,
        platform: e.platform,
        referrerId: e.eventType === 'share' ? (e.userId || e.referrerId) : e.referrerId,
        userId: e.userId,
        user: state.ui.maskPII ? (e.userEmailLower ? maskId('u', e.userEmailLower) : (e.userId ? maskId('u', e.userId) : null)) : (e.userEmail || e.userId),
        payload: e.payload
      }));
      downloadText('referral_events_filtered.json', JSON.stringify(out, null, 2), 'application/json;charset=utf-8');
    });
  }

  // -----------------------------
  // Event wiring
  // -----------------------------
  function setupUI() {
    // Drop zone / file input
    const dz = el('dropZone');
    const fi = el('fileInput');
    const btn = el('btnPickFile');
    const btnSample = el('btnLoadSample');

    dz.addEventListener('click', () => fi.click());
    btn.addEventListener('click', (e) => { e.stopPropagation(); fi.click(); });
    btnSample.addEventListener('click', (e) => { e.stopPropagation(); alert('サンプルは同梱していません。'); });

    dz.addEventListener('dragover', (e) => {
      e.preventDefault();
      dz.classList.add('dragover');
    });
    dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
    dz.addEventListener('drop', (e) => {
      e.preventDefault();
      dz.classList.remove('dragover');
      const file = e.dataTransfer.files?.[0];
      if (file) loadWorkbookFromFile(file).catch(err => {
        console.error(err);
        alert('読み込みに失敗しました（コンソールを確認してください）。');
      });
    });

    fi.addEventListener('change', () => {
      const file = fi.files?.[0];
      if (file) loadWorkbookFromFile(file).catch(err => {
        console.error(err);
        alert('読み込みに失敗しました（コンソールを確認してください）。');
      });
    });

    // toggles
    el('toggleMaskPII').addEventListener('change', () => {
      state.ui.maskPII = el('toggleMaskPII').checked;
      populateReferrerOptions();
      renderAll();
    });
    el('toggleHideUnnamed').addEventListener('change', () => {
      state.ui.hideUnnamed = el('toggleHideUnnamed').checked;
      renderDashboard();
    });

    // filters: diagnosis
    const diagInputs = ['diagDateFrom', 'diagDateTo', 'diagGender', 'diagType', 'diagAgeMin', 'diagAgeMax', 'diagReferral', 'diagUnit', 'diagAxisPick', 'diagAxisView'];
    diagInputs.forEach(id => el(id).addEventListener('change', () => renderDiagnosisTab()));
    ['diagAgeMin','diagAgeMax'].forEach(id => el(id).addEventListener('input', () => renderDiagnosisTab()));

    // favorites
    const favInputs = ['favDateFrom','favDateTo','favGender','favType','favAgeMin','favAgeMax','favReferrer','favUnit'];
    favInputs.forEach(id => el(id).addEventListener('change', () => renderFavoritesTab()));
    ['favAgeMin','favAgeMax'].forEach(id => el(id).addEventListener('input', () => renderFavoritesTab()));

    // referral
    const refInputs = ['refDateFrom','refDateTo','refEventType','refPlatform','refReferrer','refSankeyMode','refMinEdge'];
    refInputs.forEach(id => el(id).addEventListener('change', () => renderReferralTab()));
    el('refMinEdge').addEventListener('input', () => renderReferralTab());

    // scroll top
    const topBtn = el('btnScrollTop');
    window.addEventListener('scroll', () => {
      if (window.scrollY > 600) topBtn.style.display = 'block';
      else topBtn.style.display = 'none';
    });
    topBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));

    // export
    setupExportButtons();
  }

  // -----------------------------
  // Bootstrap
  // -----------------------------
  document.addEventListener('DOMContentLoaded', () => {
    setupUI();
  });

})();
