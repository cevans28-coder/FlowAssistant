/******************************************************
 * 04_state_engine.gs — State changes & Live presence
 * (v2 StatusLogs; Live stays "Live")
 *
 * Depends on:
 * - TZ, SHEETS, STATES (00_constants.gs)
 * - indexMap_, normId_, toISODate_, minutesBetween_, readRows_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_(), getSessionTokenFor_(),
 * appendStatusLogSmart_(), reconcileTimeInStateForToday_() (03_sessions.gs)
 * - updateLiveKPIsFor_(), computeLiveProductionToday_() (07_metrics.gs)
 ******************************************************/

// Explicit tabs (Live stays "Live"; state logs are v2)
var LIVE_TAB = (typeof SHEETS !== 'undefined' && SHEETS.LIVE) ? SHEETS.LIVE : 'Live';
var STATUS_LOGS_TAB = (typeof SHEETS !== 'undefined' && SHEETS.STATUS_LOGS) ? SHEETS.STATUS_LOGS : 'StatusLogs_v2';

/* ----------------- helpers ----------------- */

// Canonical Live resolver (never creates dup tabs, aligns headers)
function getLiveSheet_() {
  const HEADERS = __LIVE_HEADERS__ ? __LIVE_HEADERS__() : [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ];
  const sh = getOrCreateMasterSheet_(LIVE_TAB, HEADERS);
  const width = Math.max(HEADERS.length, sh.getLastColumn() || 1);
  const curHdr = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
  for (let i = 0; i < HEADERS.length; i++) {
    if ((curHdr[i] || '').trim() !== HEADERS[i]) sh.getRange(1, i + 1).setValue(HEADERS[i]);
  }
  sh.setFrozenRows(1);
  return sh;
}

// Optional: tiny cache buster used by Today (Quick Glance)
function _bustTodaySummaryCache_(analystId){
  try {
    const key = 'SUM:' + String(analystId||'').toLowerCase() + ':' + toISODate_(new Date());
    CacheService.getUserCache().remove(key);
  } catch(e) {}
}

/* -------------- compatibility wrapper -------------- */
function appendStatusLog_(ts, dateISO, analystId, state, source, note) {
  // legacy wrapper -> smart appender that closes the previous stint
  appendStatusLogSmart_(analystId, state, source || 'system', note || '', ts);
}

/* -------------- reads -------------- */
function getCurrentStateInfo_() {
  const ss = master_();
  const today = toISODate_(new Date());
  const id = getCurrentAnalystId_();

  const status = readRows_(ss.getSheetByName(STATUS_LOGS_TAB))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);

  const last = status[status.length - 1];
  if (!last) return { state: 'Idle', since_iso: null };

  return {
    state: String(last.state || 'Idle'),
    since_iso: last.ts ? last.ts.toISOString() : null
  };
}

/* -------------- user action -------------- */
function setState(token, newState, noteOpt) {
  const id = getCurrentAnalystId_();

  // Adopt or validate session (uses your 03_sessions helper)
  try {
    token = requireSessionOrAdopt_(token);
  } catch (e) {
    token = makeToken_();
    setSessionTokenFor_(id, token);
    try { heartbeat(token); } catch (e2) {}
  }

  // Write StatusLogs (smart appender from 04_state_engine / watchdog)
  // Fallback: if appendStatusLogSmart_ isn't present in your project,
  // replace this call with an explicit append into SHEETS.STATUS_LOGS.
  appendStatusLogSmart_(id, String(newState || 'Idle'), 'ui', String(noteOpt || ''), new Date());

  // Freshen Live + Today cache (best effort)
  try { refreshLiveFor_(id); } catch (e3) {}
  try { _bustTodaySummaryCache_(id); } catch (e4) {}

  return { ok: true, analyst_id: id, state: String(newState || 'Idle'), ts: new Date().toISOString() };
}

/* -------------- Live upsert (row writer with anti-bounce guard) -------------- */
function upsertLive_(analystId, patch) {
  if (!analystId) throw new Error('upsertLive_: missing analystId');
  patch = patch || {};

  // Optional override: set patch.__force = true to bypass anti-downgrade checks.
  var FORCE = !!patch.__force;

  var HEADERS = __LIVE_HEADERS__ ? __LIVE_HEADERS__() : [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ];

  var sh = getLiveSheet_(); // guarantees headers on LIVE_TAB
  var idx = indexMap_(HEADERS);

  // Find or append row for this analyst
  var targetIdNorm = normId_(analystId);
  var lastRow = sh.getLastRow();
  var rowIndex = -1;

  if (lastRow > 1) {
    var all = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (var r = 0; r < all.length; r++) {
      var id = normId_(all[r][idx['analyst_id']]);
      if (id && id === targetIdNorm) { rowIndex = r + 2; break; }
    }
  }
  if (rowIndex === -1) {
    // create a new row with safe defaults
    sh.appendRow([
      analystId, '', '', 'NO', '', '', '', 0, 0, '', '', '', 0, 0, 0, 0
    ]);
    rowIndex = sh.getLastRow();
  }

  var cur = sh.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];

  // helpers
  function num(v, fb){ var n = Number(v); return isFinite(n) ? n : (isFinite(fb) ? Number(fb) : 0); }
  function str(v, fb){ return (v == null ? String(fb || '') : String(v)); }
  function parseIso(s){ try{ var d = new Date(s); return (d instanceof Date && !isNaN(d)) ? d : null; }catch(e){ return null; } }

  // Existing state context
  var existingState = String(cur[idx['state']] || '');
  var existingSinceIso = String(cur[idx['since_iso']] || '');
  var existingSince = parseIso(existingSinceIso);

  // Incoming patch state context
  var incomingState = (patch.state != null) ? String(patch.state) : existingState;
  var incomingSinceIso = (patch.since_iso != null) ? String(patch.since_iso) : existingSinceIso;
  var incomingSince = parseIso(incomingSinceIso);

  // Determine "online" from the *incoming* state we intend to keep
  function onlineFor(state){ return (state && state.toLowerCase() !== 'loggedout') ? 'YES' : 'NO'; }

  // ---------------- Anti-bounce guard ----------------
  // Ignore *downgrades to Idle* (or blank) if we recently set a stronger state,
  // unless caller forces or we have a strictly newer since_iso.
  // Also ignore any state change with an older since_iso than what we already have.
  if (!FORCE) {
    var now = new Date();
    var RECENT_MS = 5 * 60 * 1000; // 5 minutes guard window

    // 1) If incoming has since_iso older than existing, refuse the state change
    if (existingSince && incomingSince && incomingSince.getTime() < existingSince.getTime()) {
      incomingState = existingState;
      incomingSinceIso = existingSinceIso;
      incomingSince = existingSince;
    }

    // 2) If trying to set Idle but we are non-Idle very recently, refuse
    var isDowngradeToIdle = (String(incomingState || '').toLowerCase() === 'idle') &&
                            (String(existingState || '').toLowerCase() !== 'idle');
    if (isDowngradeToIdle && existingSince && (now.getTime() - existingSince.getTime() < RECENT_MS)) {
      incomingState = existingState;
      incomingSinceIso = existingSinceIso;
      incomingSince = existingSince;
    }
  }

  // Compute final "online" after guard
  var nextOnline = onlineFor(incomingState);

  // Merge other fields (do not touch columns we’re not changing)
  var merged = {
    analyst_id: analystId,
    name: (patch.name != null) ? patch.name : cur[idx['name']],
    team: (patch.team != null) ? patch.team : cur[idx['team']],
    online: nextOnline,
    last_seen_iso: (patch.last_seen_iso != null) ? patch.last_seen_iso : (cur[idx['last_seen_iso']] || new Date().toISOString()),
    state: incomingState,
    since_iso: incomingSinceIso,
    mins_in_state: num((patch.mins_in_state != null) ? patch.mins_in_state : cur[idx['mins_in_state']], 0),
    today_checks: num((patch.today_checks != null) ? patch.today_checks : cur[idx['today_checks']], 0),
    baseline_hours: (patch.baseline_hours != null) ? patch.baseline_hours : cur[idx['baseline_hours']],
    location_today: (patch.location_today != null) ? patch.location_today : cur[idx['location_today']],
    session_token: (patch.session_token != null) ? patch.session_token : cur[idx['session_token']],
    logged_in_mins: num((patch.logged_in_mins != null) ? patch.logged_in_mins : cur[idx['logged_in_mins']], 0),
    live_efficiency_pct: num((patch.live_efficiency_pct != null) ? patch.live_efficiency_pct : cur[idx['live_efficiency_pct']], 0),
    live_utilisation_pct: num((patch.live_utilisation_pct != null) ? patch.live_utilisation_pct : cur[idx['live_utilisation_pct']], 0),
    live_throughput_per_hr: num((patch.live_throughput_per_hr != null) ? patch.live_throughput_per_hr : cur[idx['live_throughput_per_hr']], 0)
  };

  // Write back
  var out = new Array(HEADERS.length);
  out[idx['analyst_id']] = str(merged.analyst_id, cur[idx['analyst_id']]);
  out[idx['name']] = str(merged.name, '');
  out[idx['team']] = str(merged.team, '');
  out[idx['online']] = str(merged.online, 'NO');
  out[idx['last_seen_iso']] = str(merged.last_seen_iso, new Date().toISOString());
  out[idx['state']] = str(merged.state, '');
  out[idx['since_iso']] = str(merged.since_iso, '');
  out[idx['mins_in_state']] = num(merged.mins_in_state, 0);
  out[idx['today_checks']] = num(merged.today_checks, 0);
  out[idx['baseline_hours']] = merged.baseline_hours;
  out[idx['location_today']] = str(merged.location_today, '');
  out[idx['session_token']] = str(merged.session_token, '');
  out[idx['logged_in_mins']] = num(merged.logged_in_mins, 0);
  out[idx['live_efficiency_pct']] = num(merged.live_efficiency_pct, 0);
  out[idx['live_utilisation_pct']] = num(merged.live_utilisation_pct, 0);
  out[idx['live_throughput_per_hr']] = num(merged.live_throughput_per_hr, 0);

  sh.getRange(rowIndex, 1, 1, HEADERS.length).setValues([out]);
}

/* -------------- Live refresher (reads + calls upsert) -------------- */
function refreshLiveFor_(id) {
  const ss = master_();
  const myId = normId_(id || getCurrentAnalystId_());

  const sh = getLiveSheet_(); // ensures headers and returns the Live tab

  // Analyst profile
  const aSh = ss.getSheetByName(SHEETS.ANALYSTS);
  const aVals = aSh ? aSh.getDataRange().getValues() : [];
  const aIdx = aVals.length ? indexMap_(aVals[0].map(String)) : {};
  let name = '', team = '', baseline = 8.5;
  for (let r = 1; r < aVals.length; r++) {
    if (normId_(aVals[r][aIdx['analyst_id']]) === myId) {
      name = String(aVals[r][aIdx['name']] || '');
      team = String(aVals[r][aIdx['team']] || '');
      baseline = Number(aVals[r][aIdx['contracted_hours']]) || 8.5;
      break;
    }
  }

  const today = toISODate_(new Date());

  // Current state from StatusLogs_v2
  const slRows = readRows_(ss.getSheetByName(STATUS_LOGS_TAB))
    .filter(r => r.date_str === today && r.analyst_id_norm === myId)
    .sort((a,b) => a.ts - b.ts);
  const last = slRows[slRows.length - 1];
  const state = last ? String(last.state || 'Idle') : 'Idle';
  const sinceIso = (last && last.ts) ? last.ts.toISOString() : null;
  const minsInState = (last && last.ts) ? minutesBetween_(last.ts, new Date()) : 0;

  // Today’s checks (v2)
  const ceRows = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
    .filter(r => r.date_str === today && r.analyst_id_norm === myId);
  const todayChecks = ceRows.length;

  // Location
  const locToday = getLocationToday_(myId) || '';

  // Last-seen + online
  const up = PropertiesService.getUserProperties();
  let lastSeenIso = up.getProperty('last_seen_iso') || new Date().toISOString();
  const liveVals = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : null;
  if (liveVals && liveVals.length > 1) {
    const L = indexMap_(liveVals[0].map(String));
    for (let r = 1; r < liveVals.length; r++) {
      if (normId_(liveVals[r][L['analyst_id']]) === myId) {
        const existing = String(liveVals[r][L['last_seen_iso']] || '');
        if (existing) lastSeenIso = existing;
        break;
      }
    }
  }
  const online = minutesBetween_(new Date(lastSeenIso), new Date()) <= 5 ? 'YES' : 'NO';

  // Session token (named locally; never reference a bare "token")
  const sessToken = getSessionTokenFor_(myId);

  // Live KPIs
  const k = computeLiveProductionToday_(myId);
  const loggedInMins = Number(k.logged_in_mins) || 0;
  const liveEfficiency = Number(k.live_efficiency_pct) || 0;
  const liveUtil = Number(k.live_utilisation_pct) || 0;
  const liveTPH = Number(k.live_throughput_per_hr) || 0;

  // Upsert
  upsertLive_(myId, {
    analyst_id: myId,
    name, team,
    online,
    last_seen_iso: lastSeenIso,
    state,
    since_iso: sinceIso,
    mins_in_state: minsInState,
    today_checks: todayChecks,
    baseline_hours: baseline,
    location_today: locToday,
    session_token: sessToken,
    logged_in_mins: loggedInMins,
    live_efficiency_pct: liveEfficiency,
    live_utilisation_pct: liveUtil,
    live_throughput_per_hr: liveTPH
  });

  // keep Today cache fresh (optional)
  _bustTodaySummaryCache_(myId);
}

/* -------------- one-off repair -------------- */
function ensureLiveHeadersAndBackfill_() {
  const HEADERS = __LIVE_HEADERS__ ? __LIVE_HEADERS__() : [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ];
  const sh = getOrCreateMasterSheet_(LIVE_TAB, HEADERS);

  const current = sh.getRange(1, 1, 1, Math.max(HEADERS.length, sh.getLastColumn()))
    .getValues()[0].map(String);
  HEADERS.forEach((h, i) => {
    if ((current[i] || '').trim() !== h) sh.getRange(1, i + 1).setValue(h);
  });
  sh.setFrozenRows(1);

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

function runEnsureLiveHeadersAndBackfill() {
  ensureLiveHeadersAndBackfill_();
}

/* --------------------- Diagnostics / Smokes --------------------- */
function FA_Smoke() {
  const id = getCurrentAnalystId_();
  const out = { id, ok: true, steps: [] };

  function step(name, fn){
    try { const r = fn(); out.steps.push({name, ok:true, r}); }
    catch(e){ out.ok=false; out.steps.push({name, ok:false, err: String(e && e.message || e)}); }
  }

  step('getLiveSheet_', () => { const sh = getLiveSheet_(); return { sheet: sh.getName(), rows: sh.getLastRow() }; });
  step('getSessionTokenFor_', () => ({ token: getSessionTokenFor_(id) }));
  step('refreshLiveFor_', () => { refreshLiveFor_(id); return 'ok'; });
  step('_bustTodaySummaryCache_', () => { if (typeof _bustTodaySummaryCache_ === 'function') _bustTodaySummaryCache_(id); return 'ok'; });
  step('getTodaySummary', () => getTodaySummary && getTodaySummary());
  return out;
}

function DTS_Smoke_Me_Today(){
  const me = getCurrentAnalystId_();
  const today = toISODate_(new Date());
  const res = (typeof upsertDailyTypeSummaryFor_ === 'function')
    ? upsertDailyTypeSummaryFor_(me, today)
    : { ok:false, note:'upsertDailyTypeSummaryFor_ not found' };
  Logger.log("Smoke test result: " + JSON.stringify(res));
  SpreadsheetApp.getActive().toast("Smoke test: " + JSON.stringify(res));
  return res;
}

function SMOKE_Heartbeat() {
  // simulate UI: make sure you are “registered” so you have a session token
  try { registerMeTakeover('', ''); } catch(e) {} // safe if already registered
  const tok = getSessionTokenFor_(getCurrentAnalystId_());
  Logger.log('tok=' + tok);
  heartbeat(tok); // this calls refreshLiveFor_ internally
  Logger.log('OK heartbeat');
}

function SMOKE_SetStateWorking() {
  const tok = getSessionTokenFor_(getCurrentAnalystId_());
  setState(tok, 'Working', 'Smoke switch');
}

function SMOKE_FixBounce() {
  const id = getCurrentAnalystId_();
  const tok = getSessionTokenFor_(id); // or makeToken_ + setSessionTokenFor_
  setState(tok, 'Working', 'bounce test');
  const live = master_().getSheetByName(LIVE_TAB).getDataRange().getValues();
  Logger.log(live.map(r=>r.join('|')).join('\n'));
}

function SMOKE_StateTrace_Today(){
  const ss = master_();
  const id = getCurrentAnalystId_();
  const today = toISODate_(new Date());

  const live = ss.getSheetByName(LIVE_TAB);
  const sl = ss.getSheetByName(STATUS_LOGS_TAB);

  var liveRow = null;
  if (live && live.getLastRow()>1){
    const v = live.getDataRange().getValues();
    const L = indexMap_(v[0].map(String));
    for (let r=1;r<v.length;r++){
      if (normId_(v[r][L['analyst_id']])===id){ liveRow = v[r]; break; }
    }
  }

  const logs = readRows_(sl)
    .filter(r => r.date_str===today && r.analyst_id_norm===id)
    .sort((a,b)=> a.ts-b.ts)
    .map(r => ({ts:r.ts && r.ts.toISOString(), state:r.state, source:r.source, note:r.note}));

  Logger.log('LIVE row: ' + JSON.stringify(liveRow));
  Logger.log('StatusLogs today: ' + JSON.stringify(logs, null, 2));
  return { live_row: liveRow, logs };
}

function DEBUG_LiveTrace() {
  var ss = master_();
  var sh = ss.getSheetByName(LIVE_TAB);
  var v = sh.getDataRange().getValues();
  Logger.log(v.map(r => r.join(' | ')).join('\n'));
}
