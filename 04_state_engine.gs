/******************************************************
 * 04_state_engine.gs — State changes & Live presence
 * Depends on:
 * - TZ, SHEETS, STATES (00_constants.gs)
 * - indexMap_, normId_, toISODate_, minutesBetween_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), readRows_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_() (03_sessions.gs)
 ******************************************************/

/**
 * Append one audited status row (StatusLogs).
 */
function appendStatusLog_(ts, dateISO, analystId, state, source, note) {
  const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS,
    ['timestamp_iso','date','analyst_id','state','source','note']);
  sh.appendRow([ts.toISOString(), dateISO, analystId, state, source || 'system', note || '']);
}

/**
 * Return the current state + since for the signed-in user (today only).
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
 * User action: set state (validates against STATES).
 * Writes StatusLogs and refreshes Live.
 */
function setState(token, state, note) {
  requireSession_(token);

  if (STATES.indexOf(state) === -1) {
    throw new Error('Invalid state: ' + state);
  }

  const ts = new Date();
  const id = getCurrentAnalystId_();
  const date = toISODate_(ts);

  appendStatusLog_(ts, date, id, state, 'UI', note || '');
  refreshLiveFor_(id);
try { updateLiveKPIsFor_(id); } catch(e) {}

  return { ok: true, ts: ts.toISOString(), state, analyst_id: id };
}

/**
 * Upsert the LIVE row for an analyst.
 * - Derives online=YES unless state is LoggedOut
 * - Keeps mins_in_state fresh
 * - Mirrors today_checks & baseline_hours for convenience
 */
function upsertLive_(analystId, patch) {
  const ss = master_();
  // Add the four NEW live columns at the end:
  const sh = getOrCreateMasterSheet_(SHEETS.LIVE, [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ]);

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);

  // find existing row
  let rowIndex = -1;
  for (let r = 1; r < vals.length; r++) {
    if (normId_(vals[r][idx['analyst_id']]) === normId_(analystId)) { rowIndex = r; break; }
  }
  if (rowIndex === -1) {
    // create a new blank row with all columns so later sets have cells to write to
    sh.appendRow([
      analystId, '', '', // analyst_id, name, team
      '', '', '', '', // online, last_seen_iso, state, since_iso
      0, 0, '', '', // mins_in_state, today_checks, baseline_hours, location_today
      '', // session_token
      0, 0, 0, 0 // NEW: logged_in_mins, live_efficiency_pct, live_utilisation_pct, live_throughput_per_hr
    ]);
    rowIndex = sh.getLastRow() - 1;
  }

  const curRow = sh.getRange(rowIndex + 1, 1, 1, sh.getLastColumn()).getValues()[0];

  // derive state/online
  const currentState = String(curRow[idx['state']] || '');
  const nextState = (patch.state != null ? patch.state : currentState);
  const online = (nextState && String(nextState).toLowerCase() !== 'loggedout') ? 'YES' : 'NO';

  // build writes (only write keys that exist in header)
  const writes = {
  analyst_id: analystId,
  name: patch.name ?? curRow[idx['name']],
  team: patch.team ?? curRow[idx['team']],
  online,
  last_seen_iso: patch.last_seen_iso ?? new Date().toISOString(),
  state: nextState,
  since_iso: patch.since_iso ?? curRow[idx['since_iso']],
  mins_in_state: patch.mins_in_state ?? curRow[idx['mins_in_state']],
  today_checks: patch.today_checks ?? curRow[idx['today_checks']],
  baseline_hours: patch.baseline_hours ?? curRow[idx['baseline_hours']],
  location_today: patch.location_today ?? curRow[idx['location_today']],
  logged_in_mins: patch.logged_in_mins ?? curRow[idx['logged_in_mins']],
  live_efficiency_pct: patch.live_efficiency_pct ?? curRow[idx['live_efficiency_pct']],
  live_utilisation_pct: patch.live_utilisation_pct ?? curRow[idx['live_utilisation_pct']],
  live_throughput_per_hr: patch.live_throughput_per_hr ?? curRow[idx['live_throughput_per_hr']],
  session_token: (typeof patch.session_token !== 'undefined') ? patch.session_token : curRow[idx['session_token']]
};

  Object.keys(writes).forEach(k => {
    if (!(k in idx)) return;
    const val = writes[k];
    if (typeof val === 'undefined') return;
    sh.getRange(rowIndex + 1, idx[k] + 1).setValue(val);
  });
}

/**
 * Re-calculate and push the LIVE snapshot for an analyst.
 * - Pulls name/team/baseline from Analysts
 * - Determines today’s latest state/since from StatusLogs
 * - Counts today’s checks
 * - Derives online via last heartbeat (≤5m)
 * - Includes session_token pass-through (for forced logout compatibility)
 */
function refreshLiveFor_(id){
  const ss = master_();
  getOrCreateMasterSheet_(SHEETS.LIVE, [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today','session_token',
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ]);

  // Analyst profile
  const aSh = ss.getSheetByName(SHEETS.ANALYSTS); const aVals = aSh ? aSh.getDataRange().getValues() : [];
  const aIdx = aVals.length ? indexMap_(aVals[0]) : {};
  let name='', team='', baseline=8.5;
  for (let r=1;r<aVals.length;r++){
    if (normId_(aVals[r][aIdx['analyst_id']])===id){
      name = String(aVals[r][aIdx['name']]||'');
      team = String(aVals[r][aIdx['team']]||'');
      baseline = Number(aVals[r][aIdx['contracted_hours']])||8.5;
      break;
    }
  }

  const today = toISODate_(new Date());

  // Current state (last of today)
  const status = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r=> r.date_str===today && r.analyst_id_norm===id)
    .sort((a,b)=> a.ts-b.ts);
  const last = status[status.length-1];
  const state = last ? String(last.state) : 'Idle';
  const sinceIso = last && last.ts ? last.ts.toISOString() : null;
  const minsInState = last && last.ts ? minutesBetween_(last.ts, new Date()) : 0;

  // Checks count (today)
  const checks = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
    .filter(r=> r.date_str===today && r.analyst_id_norm===id);
  const todayChecks = checks.length;

  // Location
  const locToday = getLocationToday_(id) || '';

  // Online/last seen (from user props heartbeat)
  const up = PropertiesService.getUserProperties();
  const lastSeenIso = up.getProperty('last_seen_iso') || new Date().toISOString();
  const online = minutesBetween_(new Date(lastSeenIso), new Date()) <= 5 ? 'YES' : 'NO';
  const token = getSessionTokenFor_(id);

  // --- NEW live minutes & KPIs ---
  const loggedInMins = computeLoggedInMinutesToday_(id); // all states except LoggedOut
  const prod = computeLiveProductionToday_(id); // handling/output/standard
  const handling = prod.handling_mins;
  const standard = prod.standard_mins;

  // Live KPIs as requested
  const liveEfficiency = standard > 0 ? Math.round((handling / standard) * 100) : 0;
  const liveUtilisation = loggedInMins > 0 ? Math.round((handling / loggedInMins) * 100) : 0;
  const liveTPH = loggedInMins > 0 ? Number((checks.length / (loggedInMins/60)).toFixed(2)) : 0;

  // Upsert Live row
  upsertLive_(id, {
    analyst_id:id, name, team, online, last_seen_iso:lastSeenIso, state, since_iso:sinceIso,
    mins_in_state:minsInState, today_checks:todayChecks, baseline_hours:baseline,
    location_today: locToday, session_token:token,
    // NEW fields:
    logged_in_mins: loggedInMins,
    live_efficiency_pct: liveEfficiency,
    live_utilisation_pct: liveUtilisation,
    live_throughput_per_hr: liveTPH
  });
}
    
    
function runEnsureLiveHeadersAndBackfill() {ensureLiveHeadersAndBackfill_(); }



function ensureLiveHeadersAndBackfill_() {
  const ss = master_();
  const HEADERS = [
    'analyst_id','name','team','online','last_seen_iso','state','since_iso',
    'mins_in_state','today_checks','baseline_hours','location_today',
    'session_token',
    // KPI columns (must exist!)
    'logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'
  ];
  const sh = getOrCreateMasterSheet_(SHEETS.LIVE, HEADERS);

  // Ensure header order/columns exist
  const current = sh.getRange(1,1,1,Math.max(HEADERS.length, sh.getLastColumn())).getValues()[0].map(String);
  HEADERS.forEach((h,i) => { if ((current[i]||'').trim() !== h) sh.getRange(1,i+1).setValue(h); });
  sh.setFrozenRows(1);

  // Backfill blanks with zeros for KPI columns to avoid NaNs
  const idx = indexMap_(HEADERS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  const data = sh.getRange(2,1,lastRow-1,HEADERS.length).getValues();
  const kpiCols = ['logged_in_mins','live_efficiency_pct','live_utilisation_pct','live_throughput_per_hr'].map(k => idx[k]);
  for (let r=0;r<data.length;r++){
    kpiCols.forEach(c => {
      if (c>=0 && (data[r][c] === '' || data[r][c] == null)) data[r][c] = 0;
    });
  }
  sh.getRange(2,1,lastRow-1,HEADERS.length).setValues(data);
}
