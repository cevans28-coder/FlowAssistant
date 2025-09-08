/******************************************************
 * 04_state_engine.gs — State changes & Live presence
 * Depends on:
 * - TZ, SHEETS, STATES (00_constants.gs)
 * - indexMap_, normId_, toISODate_, minutesBetween_, readRows_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_(), getSessionTokenFor_() (03_sessions.gs)
 * - computeLoggedInMinutesToday_() (01_utils.gs)
 * - updateLiveKPIsFor_() (07_metrics.gs)
 ******************************************************/

/**
 * Append one audited status row (StatusLogs).
 * Small, hot path helper used by setState and TL actions.
 */
function appendStatusLog_(ts, dateISO, analystId, state, source, note) {
  const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS,
    ['timestamp_iso','date','analyst_id','state','source','note']);
  sh.appendRow([ts.toISOString(), dateISO, analystId, state, source || 'system', note || '']);
}

/**
 * Return the current state + since for the signed-in user (today only).
 * If no state today, default to Idle (since=null).
 */
function getCurrentStateInfo_() {
  const ss = master_();
  const today = toISODate_(new Date());
  const id = getCurrentAnalystId_();

  const status = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);

  const last = status[status.length - 1];
  if (!last) return { state: 'Idle', since_iso: null };

  return {
    state: String(last.state || 'Idle'),
    since_iso: last.ts ? last.ts.toISOString() : null
  };
}

/**
 * User action: set state (validates against STATES), logs audit entry,
 * refreshes Live snapshot, and nudges Live KPIs best-effort.
 */
function setState(token, state, note) {
  const lock = LockService.getScriptLock();
  // fail fast if another write is in progress
  if (!lock.tryLock(3000)) throw new Error('Please try again (another update is in progress).');
  try {
    requireSession_(token);
    if (STATES.indexOf(state) === -1) throw new Error('Invalid state: ' + state);

    const ts = new Date();
    const id = getCurrentAnalystId_();
    const date = toISODate_(ts);

    appendStatusLog_(ts, date, id, state, 'UI', note || '');
    refreshLiveFor_(id);
    try { updateLiveKPIsFor_(id); } catch (e) {}

    return { ok: true, ts: ts.toISOString(), state, analyst_id: id };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}


/**
 * Upsert the LIVE row for an analyst safely (header-guarded, no duplicates).
 * @param {string} analystId Normalised email/id of the analyst
 * @param {Object} patch Partial fields to update. Known keys:
 * name, team, state, since_iso, last_seen_iso, mins_in_state, today_checks,
 * baseline_hours, location_today, session_token,
 * logged_in_mins, live_efficiency_pct, live_utilisation_pct, live_throughput_per_hr
 */
function upsertLive_(analystId, patch) {
  if (!analystId) throw new Error('upsertLive_: missing analystId');
  patch = patch || {};

  const ss = master_();

  // Canonical Live headers (MUST match exactly, in order)
  const HEADERS = [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ];

  // Ensure sheet exists & header row is exact
  const sh = getOrCreateMasterSheet_(SHEETS.LIVE, HEADERS);
  const width = Math.max(HEADERS.length, sh.getLastColumn() || 1);
  const curHdr = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
  for (let i = 0; i < HEADERS.length; i++) {
    if ((curHdr[i] || '').trim() !== HEADERS[i]) {
      sh.getRange(1, i + 1).setValue(HEADERS[i]);
    }
  }
  sh.setFrozenRows(1);

  // Build a fast name → index map from the (now) canonical header row
  const idx = indexMap_(HEADERS);

  // Find an existing row for this analyst
  const targetIdNorm = normId_(analystId);
  const lastRow = sh.getLastRow();
  let rowIndex = -1;

  if (lastRow > 1) {
    const all = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (let r = 0; r < all.length; r++) {
      const id = normId_(all[r][idx['analyst_id']]);
      if (id && id === targetIdNorm) { rowIndex = r + 2; break; } // sheet row number (1-based)
    }
  }

  // If no row exists, append a new blank one with defaults
  if (rowIndex === -1) {
    const blank = [
      analystId, // analyst_id
      '', // name
      '', // team
      'NO', // online
      '', // last_seen_iso
      '', // state
      '', // since_iso
      0, // mins_in_state
      0, // today_checks
      '', // baseline_hours (kept as number/string; not critical)
      '', // location_today
      '', // session_token
      0, // logged_in_mins
      0, // live_efficiency_pct
      0, // live_utilisation_pct
      0 // live_throughput_per_hr
    ];
    sh.appendRow(blank);
    rowIndex = sh.getLastRow();
  }

  // Read current row as the base to merge with
  const cur = sh.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];

  // Derive next state and online flag
  const currentState = String(cur[idx['state']] || '');
  const nextState = (patch.state != null ? String(patch.state) : currentState);
  const online = (nextState && nextState.toLowerCase() !== 'loggedout') ? 'YES' : 'NO';

  // Helper: numeric-safe
  const num = (v, fallback) => {
    const n = Number(v);
    return isFinite(n) ? n : Number(fallback) || 0;
  };
  const str = (v, fallback) => (v == null ? String(fallback || '') : String(v));

  // Merge current row with patch (write only known columns)
  // NOTE: last_seen_iso defaults to "now" if patch didn’t provide anything
  const nowIso = new Date().toISOString();
  const merged = {
    analyst_id: analystId,
    name: patch.name ?? cur[idx['name']],
    team: patch.team ?? cur[idx['team']],
    online: online,
    last_seen_iso: patch.last_seen_iso ?? (cur[idx['last_seen_iso']] || nowIso),
    state: nextState,
    since_iso: patch.since_iso ?? cur[idx['since_iso']],
    mins_in_state: num(patch.mins_in_state ?? cur[idx['mins_in_state']], 0),
    today_checks: num(patch.today_checks ?? cur[idx['today_checks']], 0),
    baseline_hours: patch.baseline_hours ?? cur[idx['baseline_hours']],
    location_today: patch.location_today ?? cur[idx['location_today']],
    session_token: (typeof patch.session_token !== 'undefined') ? patch.session_token : cur[idx['session_token']],
    logged_in_mins: num(patch.logged_in_mins ?? cur[idx['logged_in_mins']], 0),
    live_efficiency_pct: num(patch.live_efficiency_pct ?? cur[idx['live_efficiency_pct']], 0),
    live_utilisation_pct: num(patch.live_utilisation_pct ?? cur[idx['live_utilisation_pct']], 0),
    live_throughput_per_hr:num(patch.live_throughput_per_hr ?? cur[idx['live_throughput_per_hr']], 0)
  };

  // Build a full row buffer in header order
  const out = new Array(HEADERS.length);
  out[idx['analyst_id']] = str(merged.analyst_id, cur[idx['analyst_id']]);
  out[idx['name']] = str(merged.name, '');
  out[idx['team']] = str(merged.team, '');
  out[idx['online']] = str(merged.online, 'NO');
  out[idx['last_seen_iso']] = str(merged.last_seen_iso, nowIso);
  out[idx['state']] = str(merged.state, '');
  out[idx['since_iso']] = str(merged.since_iso, '');
  out[idx['mins_in_state']] = num(merged.mins_in_state, 0);
  out[idx['today_checks']] = num(merged.today_checks, 0);
  out[idx['baseline_hours']] = merged.baseline_hours; // keep as-is (number or string ok)
  out[idx['location_today']] = str(merged.location_today, '');
  out[idx['session_token']] = str(merged.session_token, '');
  out[idx['logged_in_mins']] = num(merged.logged_in_mins, 0);
  out[idx['live_efficiency_pct']] = num(merged.live_efficiency_pct, 0);
  out[idx['live_utilisation_pct']] = num(merged.live_utilisation_pct, 0);
  out[idx['live_throughput_per_hr']] = Number(merged.live_throughput_per_hr) || 0;

  // Write once
  sh.getRange(rowIndex, 1, 1, HEADERS.length).setValues([out]);
}
/**
 * Re-calculate and push the LIVE snapshot for an analyst (no dependency on computeLoggedInMinutesToday_).
 * - Pulls name/team/baseline from Analysts
 * - Determines today’s latest state/since from StatusLogs
 * - Counts today’s checks
 * - Derives online via last heartbeat (≤5m)
 * - Uses computeLiveProductionToday_ to get logged_in_mins + KPIs
 */
function refreshLiveFor_(id) {
  const ss = master_();
  // Ensure headers (includes KPI columns)
  getOrCreateMasterSheet_(SHEETS.LIVE, [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ]);

  // Analyst profile
  const aSh = ss.getSheetByName(SHEETS.ANALYSTS);
  const aVals = aSh ? aSh.getDataRange().getValues() : [];
  const aIdx = aVals.length ? indexMap_(aVals[0]) : {};
  let name = '', team = '', baseline = 8.5;
  for (let r = 1; r < aVals.length; r++) {
    if (normId_(aVals[r][aIdx['analyst_id']]) === id) {
      name = String(aVals[r][aIdx['name']] || '');
      team = String(aVals[r][aIdx['team']] || '');
      baseline = Number(aVals[r][aIdx['contracted_hours']]) || 8.5;
      break;
    }
  }

  const today = toISODate_(new Date());

  // Current state (last of today)
  const status = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);
  const last = status[status.length - 1];
  const state = last ? String(last.state) : 'Idle';
  const sinceIso = last && last.ts ? last.ts.toISOString() : null;
  const minsInState = last && last.ts ? minutesBetween_(last.ts, new Date()) : 0;

  // Checks count (today)
  const checks = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id);
  const todayChecks = checks.length;

  // Location
  const locToday = getLocationToday_(id) || '';

  // Online/last seen from heartbeat
  const up = PropertiesService.getUserProperties();
  const lastSeenIso = up.getProperty('last_seen_iso') || new Date().toISOString();
  const online = minutesBetween_(new Date(lastSeenIso), new Date()) <= 5 ? 'YES' : 'NO';
  const token = getSessionTokenFor_(id);

  // ---- Live minutes & KPIs from a single source of truth ----
  const prod = computeLiveProductionToday_(id); // returns logged_in_mins + KPIs
  const loggedInMins = Number(prod.logged_in_mins) || 0;
  const liveEfficiency = Number(prod.live_efficiency_pct) || 0;
  const liveUtilisation = Number(prod.live_utilisation_pct) || 0;
  const liveTPH = Number(prod.live_throughput_per_hr) || 0;

  // Upsert Live row
  upsertLive_(id, {
    analyst_id: id,
    name,
    team,
    online,
    last_seen_iso: lastSeenIso,
    state,
    since_iso: sinceIso,
    mins_in_state: minsInState,
    today_checks: todayChecks,
    baseline_hours: baseline,
    location_today: locToday,
    session_token: token,
    // KPIs/mins
    logged_in_mins: loggedInMins,
    live_efficiency_pct: liveEfficiency,
    live_utilisation_pct: liveUtilisation,
    live_throughput_per_hr: liveTPH
  });
}

/** Utility: one-off header fix + fill zeros in KPI columns (admin-safe). */
function ensureLiveHeadersAndBackfill_() {
  const ss = master_();
  const HEADERS = [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today',
    'session_token',
    // KPI columns
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ];
  const sh = getOrCreateMasterSheet_(SHEETS.LIVE, HEADERS);

  // Enforce header positions/order
  const current = sh.getRange(1, 1, 1, Math.max(HEADERS.length, sh.getLastColumn()))
    .getValues()[0].map(String);
  HEADERS.forEach((h, i) => {
    if ((current[i] || '').trim() !== h) sh.getRange(1, i + 1).setValue(h);
  });
  sh.setFrozenRows(1);

  // Backfill KPI blanks with zeros to avoid #NUM!/NaN in consumers
  const idx = indexMap_(HEADERS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  const data = sh.getRange(2, 1, lastRow - 1, Math.max(HEADERS.length, sh.getLastColumn())).getValues();
  const kpiCols = ['logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr']
    .map(k => idx[k]).filter(c => c != null);

  for (let r = 0; r < data.length; r++) {
    kpiCols.forEach(c => {
      if (c >= 0 && (data[r][c] === '' || data[r][c] == null)) data[r][c] = 0;
    });
  }
  sh.getRange(2, 1, data.length, data[0].length).setValues(data);
}

/** Handy wrapper you can run once from the editor to repair LIVE headers/data. */
function runEnsureLiveHeadersAndBackfill() {
  ensureLiveHeadersAndBackfill_();
}

/**
 * Re-calculate and upsert the LIVE row for an analyst.
 * - Recomputes: name, team, baseline, current state/since, today_checks
 * - Derives online/last_seen from heartbeat
 * - Computes live KPIs via computeLiveProductionToday_
 * - ALWAYS writes today_checks so it can’t silently drift
 */
function refreshLiveFor_(id){
  const ss = master_();

  // Ensure LIVE has the canonical headers before we touch it
  const sh = getOrCreateMasterSheet_(SHEETS.LIVE, [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ]);

  // --- Analyst profile
  const aSh = ss.getSheetByName(SHEETS.ANALYSTS);
  const aVals = aSh ? aSh.getDataRange().getValues() : [];
  const aIdx = aVals.length ? indexMap_(aVals[0].map(String)) : {};
  let name = '', team = '', baseline = 8.5;
  for (let r = 1; r < aVals.length; r++) {
    if (normId_(aVals[r][aIdx['analyst_id']]) === id) {
      name = String(aVals[r][aIdx['name']] || '');
      team = String(aVals[r][aIdx['team']] || '');
      baseline = Number(aVals[r][aIdx['contracted_hours']]) || 8.5;
      break;
    }
  }

  const today = toISODate_(new Date());

  // --- Current state (last of today)
  const status = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a,b) => a.ts - b.ts);
  const last = status[status.length - 1];
  const state = last ? String(last.state || 'Idle') : 'Idle';
  const sinceIso = last && last.ts ? last.ts.toISOString() : null;
  const minsInState = last && last.ts ? minutesBetween_(last.ts, new Date()) : 0;

  // --- Today checks (authoritative from CheckEvents)
  const ceToday = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id);
  const todayChecks = ceToday.length;

  // --- Location
  const locToday = getLocationToday_(id) || '';

  // --- Online/last seen
  const up = PropertiesService.getUserProperties();
  const lastSeenIso = up.getProperty('last_seen_iso') || new Date().toISOString();
  const online = minutesBetween_(new Date(lastSeenIso), new Date()) <= 5 ? 'YES' : 'NO';
  const token = getSessionTokenFor_(id);

  // --- Live KPIs
  const k = computeLiveProductionToday_(id); // must exist in your project
  const loggedInMins = k.logged_in_mins || 0;
  const liveEfficiency = k.live_efficiency_pct || 0;
  const liveUtil = k.live_utilisation_pct || 0;
  const liveTPH = k.live_throughput_per_hr || 0;

  // --- Upsert the Live row (explicitly set today_checks every time)
  upsertLive_(id, {
    analyst_id: id,
    name, team,
    online,
    last_seen_iso: lastSeenIso,
    state,
    since_iso: sinceIso,
    mins_in_state: minsInState,
    today_checks: todayChecks, // <- always written
    baseline_hours: baseline,
    location_today: locToday,
    session_token: token,
    logged_in_mins: loggedInMins,
    live_efficiency_pct: liveEfficiency,
    live_utilisation_pct: liveUtil,
    live_throughput_per_hr: liveTPH
  });
}
