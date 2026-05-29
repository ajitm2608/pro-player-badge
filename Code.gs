/**
 * Pro Player Badge — JSON API backed by a Google Sheet.
 *
 * Deploy: Extensions → Apps Script → paste this file → Deploy → Manage deployments
 *         → edit → New version → Deploy. (Execute as: Me · Who has access: Anyone)
 *
 * Endpoints (GET):
 *   ?mobile=9876543210   → lookup by mobile_no; response BUNDLES cohort stats
 *   ?host_id=59682201    → lookup by host_id; response BUNDLES cohort stats
 *   ?cohort=1            → cohort stats only { ok, stats:{tasks_done_0..3,...} }
 *   ?leaderboard=1       → full ranked list (mobiles masked)
 *   ?ping=1              → health check + tab names
 *
 * v2 (May 2026): the mobile/host lookup now includes `cohort` in its payload so
 * the frontend needs only ONE round-trip on the host-lookup path instead of two.
 */

const CONFIG = {
  SHEET_NAME: 'data',              // EDIT to your data tab name (the tab at gid=870500265)
  COHORT_SHEET_NAME: 'cohort_stats', // key/value tab: col A = metric_key, col B = number
  LOG_SHEET_NAME: 'event_log',       // append-only log: col A = timestamp, B = mobile, C = event, D = source

  COLUMNS: {
    name: 'name',
    mobile_no: 'mobile_no',
    host_id: 'host_id',
    total_games_played: 'total_games_played',
    completed_games: 'completed_games',
    repeat_users: 'repeat_users',
    total_users: 'total_users',
  },

  THRESHOLDS: {
    ASPIRING_GAMES: 25,
    PRO_GAMES: 100,
    FANS_BANAO: 15,
    REACH_BADAO: 50,
    PURA_KHELO: 0.5,
  },

  CACHE_TTL_SECONDS: 300,
  COHORT_CACHE_TTL_SECONDS: 1800,
};

function doGet(e) {
  try {
    const params = e && e.parameter || {};

    if (params.ping) {
      return jsonResponse(healthCheck_());
    }

    // Lightweight event log: ?log=<event>&mobile=10digits
    // Side effect: appends a row to the event_log tab. Fire-and-forget from the
    // client; we always return ok:true so the client doesn't retry.
    if (params.log) {
      const m = normalizeMobile_(params.mobile);
      const ev = String(params.log).trim().slice(0, 60);
      if (m && ev) {
        logEvent_(m, ev, {
          name:    params.name    || '',
          host_id: params.host_id || '',
          tier:    params.tier    || '',
          games:   params.games   || '',
          repeats: params.repeats || '',
          reach:   params.reach   || '',
        });
      }
      return jsonResponse({ ok: true, logged: !!(m && ev) });
    }

    if (params.leaderboard) {
      const cache = CacheService.getScriptCache();
      const cached = cache.get('leaderboard:v1');
      if (cached) return jsonResponse(JSON.parse(cached));
      const payload = getLeaderboard_();
      cache.put('leaderboard:v1', JSON.stringify(payload), CONFIG.CACHE_TTL_SECONDS);
      return jsonResponse(payload);
    }

    if (params.cohort) {
      return jsonResponse({ ok: true, stats: readCohortCached_() });
    }

    const mobile = normalizeMobile_(params.mobile);
    const hostId = (params.host_id || '').toString().trim();

    if (!mobile && !hostId) {
      return jsonResponse({ ok: false, error: 'Pass ?mobile=10digits, ?host_id=N, ?leaderboard=1, ?cohort=1, or ?ping=1' }, 400);
    }

    const cacheKey = mobile ? 'm:' + mobile : 'h:' + hostId;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return jsonResponse(JSON.parse(cached));
    }

    const cohort = readCohortCached_(); // bundle cohort → one round-trip for the client

    const row = findRow_(mobile, hostId);
    if (!row) {
      const payload = { ok: true, found: false, mobile: mobile, host_id: hostId, cohort: cohort };
      cache.put(cacheKey, JSON.stringify(payload), 60);
      return jsonResponse(payload);
    }

    const payload = { ok: true, found: true, cohort: cohort, ...computeBadge_(row) };
    cache.put(cacheKey, JSON.stringify(payload), CONFIG.CACHE_TTL_SECONDS);
    return jsonResponse(payload);
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) }, 500);
  }
}

function healthCheck_() {
  const ss = SpreadsheetApp.getActive();
  const tabs = ss.getSheets().map(function (s) {
    try { return { name: s.getName(), rows: s.getLastRow() }; }
    catch (e) { return { name: s.getName(), rows: null, note: 'unreadable (datasource?)' }; }
  });
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  return {
    ok: true,
    spreadsheet: ss.getName(),
    configured_sheet: CONFIG.SHEET_NAME,
    configured_sheet_found: !!sheet,
    all_tabs: tabs,
  };
}

function findRow_(mobile, hostId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Tab not found: ' + CONFIG.SHEET_NAME + '. Hit ?ping=1 to list tabs.');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const headers = data[0].map(function (h) { return String(h).trim(); });
  const idx = {};
  Object.keys(CONFIG.COLUMNS).forEach(function (k) { idx[k] = headers.indexOf(CONFIG.COLUMNS[k]); });

  const missing = Object.keys(idx).filter(function (k) { return idx[k] === -1; });
  if (missing.length) {
    throw new Error('Missing columns: ' + missing.map(function (k) { return CONFIG.COLUMNS[k]; }).join(', '));
  }

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (mobile) {
      const cellMobile = normalizeMobile_(r[idx.mobile_no]);
      if (cellMobile && cellMobile === mobile) return rowToObject_(r, idx);
    }
    if (hostId) {
      if (String(r[idx.host_id]).trim() === hostId) return rowToObject_(r, idx);
    }
  }
  return null;
}

function rowToObject_(row, idx) {
  const o = {};
  Object.keys(idx).forEach(function (k) { o[k] = row[idx[k]]; });
  return o;
}

function computeBadge_(row) {
  const totalGames = toInt_(row.total_games_played);
  const completedGames = toInt_(row.completed_games);
  const repeatUsers = toInt_(row.repeat_users);
  const totalUsers = toInt_(row.total_users);
  const completionRate = totalGames > 0 ? completedGames / totalGames : 0;

  let tier;
  if (completedGames >= CONFIG.THRESHOLDS.PRO_GAMES) tier = 'pro';
  else if (completedGames >= CONFIG.THRESHOLDS.ASPIRING_GAMES) tier = 'aspiring';
  else tier = 'start';

  const tasks = {
    fans_banao: { done: repeatUsers >= CONFIG.THRESHOLDS.FANS_BANAO, progress: repeatUsers, target: CONFIG.THRESHOLDS.FANS_BANAO },
    reach_badao: { done: totalUsers >= CONFIG.THRESHOLDS.REACH_BADAO, progress: totalUsers, target: CONFIG.THRESHOLDS.REACH_BADAO },
    pura_khelo: { done: completionRate >= CONFIG.THRESHOLDS.PURA_KHELO, progress: Math.round(completionRate * 1000) / 10, target: CONFIG.THRESHOLDS.PURA_KHELO * 100 },
  };
  const tasksDone = ['fans_banao', 'reach_badao', 'pura_khelo'].filter(function (k) { return tasks[k].done; }).length;

  let multiplier = 0;
  if (tier === 'pro') {
    multiplier = tasksDone === 3 ? 1.5 : tasksDone === 2 ? 1.2 : tasksDone === 1 ? 1.1 : 1.0;
  }

  // Mobile is intentionally NOT returned (caller already has it; avoid leaking into logs/cache).
  return {
    name: row.name,
    host_id: String(row.host_id),
    stats: {
      total_games_played: totalGames,
      completed_games: completedGames,
      repeat_users: repeatUsers,
      total_users: totalUsers,
      completion_rate_pct: Math.round(completionRate * 1000) / 10,
    },
    tier: tier,
    tasks: tasks,
    tasks_done: tasksDone,
    multiplier: multiplier,
  };
}

function normalizeMobile_(raw) {
  if (raw == null) return '';
  let s = String(raw).replace(/\D/g, '');
  if (s.length === 12 && s.indexOf('91') === 0) s = s.slice(2);
  if (s.length === 11 && s.charAt(0) === '0') s = s.slice(1);
  return s.length === 10 ? s : '';
}

function toInt_(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? 0 : n;
}

function jsonResponse(payload, _status) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getLeaderboard_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Tab not found: ' + CONFIG.SHEET_NAME);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, count: 0, hosts: [] };

  const headers = data[0].map(function (h) { return String(h).trim(); });
  const idx = {};
  Object.keys(CONFIG.COLUMNS).forEach(function (k) { idx[k] = headers.indexOf(CONFIG.COLUMNS[k]); });

  const hosts = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const games_total = toInt_(r[idx.total_games_played]);
    const games_done = toInt_(r[idx.completed_games]);
    if (games_total === 0 && games_done === 0) continue;

    const repeats = toInt_(r[idx.repeat_users]);
    const reach = toInt_(r[idx.total_users]);
    const completion_pct = games_total > 0 ? Math.round((games_done / games_total) * 1000) / 10 : 0;
    const score = games_done + repeats * 5 + reach; // placeholder formula
    const tasksDone =
      (repeats >= CONFIG.THRESHOLDS.FANS_BANAO ? 1 : 0) +
      (reach >= CONFIG.THRESHOLDS.REACH_BADAO ? 1 : 0) +
      (completion_pct >= (CONFIG.THRESHOLDS.PURA_KHELO * 100) ? 1 : 0);
    const rate = [1.0, 1.1, 1.2, 1.5][tasksDone];

    hosts.push({
      name: String(r[idx.name] || ''),
      mobile: maskMobile_(r[idx.mobile_no]),
      host_id: String(r[idx.host_id] || ''),
      games_total: games_total, games_done: games_done, completion_pct: completion_pct,
      repeats: repeats, reach: reach, score: score, rate: rate, rank: 0,
    });
  }
  hosts.sort(function (a, b) { return b.score !== a.score ? b.score - a.score : b.games_done - a.games_done; });
  hosts.forEach(function (h, i) { h.rank = i + 1; });
  return { ok: true, count: hosts.length, hosts: hosts };
}

function maskMobile_(raw) {
  const s = normalizeMobile_(raw);
  if (s.length !== 10) return '';
  return s.slice(0, 2) + '******' + s.slice(-2);
}

/**
 * Returns the cohort stats map, served from the script cache when warm so the
 * merged host-lookup path doesn't re-read the sheet on every request.
 */
function readCohortCached_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('cohort:v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }
  const stats = getCohortStats_().stats || {};
  cache.put('cohort:v1', JSON.stringify(stats), CONFIG.COHORT_CACHE_TTL_SECONDS);
  return stats;
}

function getCohortStats_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.COHORT_SHEET_NAME);
  if (!sheet) return { ok: true, stats: {}, note: 'tab "' + CONFIG.COHORT_SHEET_NAME + '" not found' };

  const data = sheet.getDataRange().getValues();
  const stats = {};
  for (let i = 0; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    const v = data[i][1];
    if (!k) continue;
    if (i === 0 && typeof v !== 'number') continue; // skip header row
    const n = Number(v);
    if (!isNaN(n)) stats[k] = n;
  }
  return { ok: true, stats: stats };
}

/**
 * Append a row to the event_log sheet tab. Used by the ?log endpoint to
 * capture user-level events (login, interest_pro_player, etc.) for the BI
 * pipeline. Failures are swallowed — logging must never break the request.
 *
 * Sheet schema (tab name from CONFIG.LOG_SHEET_NAME):
 *   col A: timestamp (Date — Sheets renders as ISO / locale string)
 *   col B: mobile_no (10 digits, normalized)
 *   col C: event     (e.g. 'login', 'interest_pro_player')
 *   col D: source    ('web' by default)
 *   col E: name      (host name at the moment of the event)
 *   col F: host_id
 *   col G: tier      ('start' | 'aspiring' | 'pro')
 *   col H: games_done
 *   col I: repeats   (subscribers count)
 *   col J: reach     (unique players)
 *
 * Cols E..J are populated when the client sends them (gate submit, auto-login,
 * interest popup). For raw `?log=…&mobile=…` calls without extras, those
 * cells stay blank — no error.
 */
function logEvent_(mobile, event, extra) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!sheet) return; // tab doesn't exist yet — skip silently
    const x = extra || {};
    sheet.appendRow([
      new Date(),
      mobile,
      event,
      'web',
      x.name    || '',
      x.host_id || '',
      x.tier    || '',
      x.games   || '',
      x.repeats || '',
      x.reach   || '',
    ]);
  } catch (e) {
    // Never throw from logging.
  }
}
