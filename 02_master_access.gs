/******************************************************
 * 03_sessions.gs — Identity & session lifecycle
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - normId_, toISODate_, minutesBetween_, indexMap_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), readRows_() (02_master_access.gs)
 * - upsertLive_(), refreshLiveFor_() (04_state_engine.gs)
 * - updateLiveKPIsFor_() (07_metrics)
 ******************************************************/

/* ===================== Identity helpers ===================== */

/** Resolve the current user's analyst_id (normalised email). */
function getCurrentAnalystId_() {
  const act = Session.getActiveUser() && Session.getActiveUser().getEmail();
  if (act && act.indexOf('@') > -1) return normId_(act);

  const eff = Session.getEffectiveUser() && Session.getEffectiveUser().getEmail();
  if (eff && eff.indexOf('@') > -1) return normId_(eff);

  const saved = PropertiesService.getUserProperties().getProperty('analyst_id');
  return normId_(saved || 'unknown_user');
}

/** Convenience: my raw email (not normalised). */
function getMyEmail_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
}

/** Generate a random session token. */
function makeToken_() {
  return Utilities.getUuid();
}

/** Ensure StatusLogs sheet + headers including time_in_state. */
function getStatusLogsSheetEnsured_() {
  const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS, [
    'timestamp_iso','date','analyst_id','state','source','note','time_in_state'
  ]);
  const v = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const need = ['timestamp_iso','date','analyst_id','state','source','note','time_in_state'];
  need.forEach((h,i)=>{ if ((v[i]||'').trim() !== h) sh.getRange(1,i+1).setValue(h); });
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Append a StatusLogs row AND close the previous stint for this analyst:
 * - Fills time_in_state on the *previous* row = (now - prev.timestamp)
 * - Appends the new row with blank time_in_state
 */
function appendStatusLogSmart_(analystId, state, source, note, tsOpt) {
  const id = normId_(analystId || getCurrentAnalystId_());
  const ts = tsOpt instanceof Date ? tsOpt : new Date();
  const dateISO = toISODate_(ts);
  const sh = getStatusLogsSheetEnsured_();
  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0].map(String));
  const A = idx['analyst_id'], T = idx['timestamp_iso'], D = idx['date'], DUR = idx['time_in_state'];

  // find last log row for this analyst
  let lastRow = -1, lastTs = 0;
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][A]) !== id) continue;
    const t = Date.parse(String(v[r][T] || ''));
    if (!isNaN(t) && t >= lastTs) { lastTs = t; lastRow = r + 1; }
  }

  // close previous stint if it has no duration yet
  if (lastRow !== -1) {
    const prevTsStr = String(sh.getRange(lastRow, T+1).getValue() || '');
    const prevTs = new Date(prevTsStr);
    if (prevTs instanceof Date && !isNaN(prevTs)) {
      const durCell = sh.getRange(lastRow, DUR+1);
      const existing = Number(durCell.getValue() || 0);
      if (!existing || existing <= 0) {
        const mins = Math.max(0, Math.round(minutesBetween_(prevTs, ts)));
        durCell.setValue(mins);
      }
    }
  }

  // append new log (duration blank; will be closed on next change)
  sh.appendRow([ts.toISOString(), dateISO, id, state, source || '', note || '', '']);
}

/**
 * Reconcile today's missing time_in_state for current (or provided) analyst.
 * Safe to call from heartbeat(); fills gaps created by external writers.
 */
function reconcileTimeInStateForToday_(analystIdOpt) {
  const id = normId_(analystIdOpt || getCurrentAnalystId_());
  const today = toISODate_(new Date());
  const sh = getStatusLogsSheetEnsured_();
  if (sh.getLastRow() < 2) return;

  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0].map(String));
  const T = idx['timestamp_iso'], D = idx['date'], A = idx['analyst_id'], DUR = idx['time_in_state'];

  const rows = [];
  for (let r = 1; r < v.length; r++) {
    if (String(v[r][D] || '') !== today) continue;
    if (normId_(v[r][A]) !== id) continue;
    const ts = new Date(String(v[r][T] || ''));
    rows.push({ r: r+1, ts });
  }
  rows.sort((x,y)=> x.ts - y.ts);
  if (!rows.length) return;

  const now = new Date();
  for (let i = 0; i < rows.length; i++) {
    const cur = rows[i];
    const next = rows[i+1];
    const cell = sh.getRange(cur.r, DUR+1);
    const has = Number(cell.getValue() || 0);
    if (has && has > 0) continue;

    const endTs = next ? next.ts : now;
    if (!(cur.ts instanceof Date) || isNaN(cur.ts) || !(endTs instanceof Date) || isNaN(endTs) || endTs <= cur.ts) continue;

    const mins = Math.max(0, Math.round(minutesBetween_(cur.ts, endTs)));
    cell.setValue(mins);
  }
}
/* ===================== Token I/O (Live) ===================== */

/** Return the most recent session token for an analyst from Live (by last_seen_iso). */
function getSessionTokenFor_(id) {
  const sh = master_().getSheetByName(SHEETS.LIVE);
  if (!sh || sh.getLastRow() < 2) return '';
  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0].map(String));

  let bestRow = null, bestTs = 0;
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) !== normId_(id)) continue;
    const tsStr = String(v[r][idx['last_seen_iso']] || '');
    const ts = Date.parse(tsStr);
    if (!isNaN(ts) && ts >= bestTs) { bestTs = ts; bestRow = v[r]; }
  }
  if (!bestRow) return '';
  return String(bestRow[idx['session_token']] || '');
}

/** Write/update session token for an analyst in the most recent Live row (creates if missing). */
function setSessionTokenFor_(id, token) {
  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.LIVE);
  const nowIso = new Date().toISOString();

  if (!sh || sh.getLastRow() < 2) {
    upsertLive_(id, { analyst_id:id, session_token:token, last_seen_iso:nowIso });
    return;
  }

  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0].map(String));

  let bestRowIdx = -1, bestTs = 0;
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) !== normId_(id)) continue;
    const ts = Date.parse(String(v[r][idx['last_seen_iso']] || ''));
    if (!isNaN(ts) && ts >= bestTs) { bestTs = ts; bestRowIdx = r + 1; }
  }

  if (bestRowIdx === -1) {
    upsertLive_(id, { analyst_id:id, session_token:token, last_seen_iso:nowIso });
    return;
  }

  if (idx['session_token'] != null) sh.getRange(bestRowIdx, idx['session_token'] + 1).setValue(token);
  if (idx['last_seen_iso'] != null) sh.getRange(bestRowIdx, idx['last_seen_iso'] + 1).setValue(nowIso);
}

/* ===================== Session control ===================== */

/**
 * Clear a user's session (token→'', online→NO, state→LoggedOut, since→now).
 * Called by logOff() and watchdog/admin actions.
 */
function clearSession_(id) {
  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.LIVE);
  if (!sh || sh.getLastRow() < 2) return;

  const vals = sh.getDataRange().getValues();
  const idx = indexMap_(vals[0].map(String));
  const aCol = idx['analyst_id'];

  let row = -1;
  for (let r = 1; r < vals.length; r++) {
    if (normId_(vals[r][aCol]) === normId_(id)) { row = r + 1; break; }
  }
  const nowIso = new Date().toISOString();
  if (row !== -1) {
    const safeSet = (name, value) => {
      if (idx[name] != null) sh.getRange(row, idx[name] + 1).setValue(value);
    };
    safeSet('session_token', '');
    safeSet('online', 'NO');
    safeSet('last_seen_iso', nowIso);
    safeSet('state', 'LoggedOut');
    safeSet('since_iso', nowIso);
  }
  PropertiesService.getUserProperties().deleteProperty('last_seen_iso');
}

/**
 * Require a valid session token for the current user.
 * - If no token is provided, allow (background calls).
 * - If token differs from Live row:
 * adopt it if the Live session looks stale (LoggedOut / offline > 5m),
 * otherwise throw to prevent concurrent sessions.
 */
function requireSession_(token) {
  const id = getCurrentAnalystId_();

  if (!token) return { ok: true, analyst_id: id };

  const current = getSessionTokenFor_(id);
  if (!current || token === current) return { ok: true, analyst_id: id };

  // Inspect the Live row to decide adoption vs. deny
  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.LIVE);
  if (sh && sh.getLastRow() > 1) {
    const v = sh.getDataRange().getValues();
    const idx = indexMap_(v[0].map(String));
    const aCol = idx['analyst_id'], sCol = idx['state'], lsCol = idx['last_seen_iso'];

    for (let r = 1; r < v.length; r++) {
      if (normId_(v[r][aCol]) !== id) continue;
      const liveState = String(v[r][sCol] || '');
      const lastSeen = new Date(String(v[r][lsCol] || new Date(0).toISOString()));
      const offline = minutesBetween_(lastSeen, new Date()) > 5;

      if (liveState === 'LoggedOut' || offline) {
        setSessionTokenFor_(id, token); // adopt this window
        return { ok: true, analyst_id: id, adopted: true };
      }
      throw new Error('You are already logged in elsewhere. Please log off on the other device first.');
    }
  }

  // No Live row case → adopt
  setSessionTokenFor_(id, token);
  return { ok: true, analyst_id: id, adopted: true };
}

/* ===================== Heartbeat / logout ===================== */

function heartbeat(token) {
  requireSession_(token);

  // Update per-user last-seen (read by refreshLiveFor_)
  PropertiesService.getUserProperties().setProperty('last_seen_iso', new Date().toISOString());

  const id = getCurrentAnalystId_();
  refreshLiveFor_(id);

  // Refresh live KPIs (best-effort)
  try { updateLiveKPIsFor_(id); } catch (e) {}

  // NEW: reconcile StatusLogs durations for today (covers external writes)
  try { reconcileTimeInStateForToday_(id); } catch (e) {}

  return { ok: true };
}

/**
 * User-initiated logout from the UI.
 * Writes an audited StatusLogs row, clears session, refreshes Live.
 * NOTE: StatusLogs now includes 'time_in_state' as the last column.
 */
function logOff(token, note) {
  try { requireSession_(token); } catch(e) { /* proceed to clean up */ }
  const id = getCurrentAnalystId_();
  const ts = new Date();

  // Smart append: closes previous stint then writes LoggedOut row
  appendStatusLogSmart_(id, 'LoggedOut', 'UI', note || 'User log off', ts);

  // collapse any duplicate Live rows first
  try { dedupeLiveRowsForAnalyst_(id); } catch(e) {}

  // clear session fields on the newest row
  const live = master_().getSheetByName(SHEETS.LIVE);
  if (live && live.getLastRow() > 1) {
    const v = live.getDataRange().getValues();
    const idx = indexMap_(v[0].map(String));
    let bestRow = -1, bestTs = 0;
    for (let r=1;r<v.length;r++){
      if (normId_(v[r][idx['analyst_id']]) !== id) continue;
      const t = Date.parse(String(v[r][idx['last_seen_iso']]||''));
      if (!isNaN(t) && t >= bestTs) { bestTs = t; bestRow = r+1; }
    }
    if (bestRow !== -1) {
      if (idx['session_token'] != null) live.getRange(bestRow, idx['session_token']+1).setValue('');
      if (idx['online'] != null) live.getRange(bestRow, idx['online']+1).setValue('NO');
      if (idx['last_seen_iso'] != null) live.getRange(bestRow, idx['last_seen_iso']+1).setValue(ts.toISOString());
      if (idx['state'] != null) live.getRange(bestRow, idx['state']+1).setValue('LoggedOut');
      if (idx['since_iso'] != null) live.getRange(bestRow, idx['since_iso']+1).setValue(ts.toISOString());
    }
  }

  PropertiesService.getUserProperties().deleteProperty('last_seen_iso');
  logLoginEvent_('Logout', note || 'User log off', token || '');
  refreshLiveFor_(id);
  return { ok: true };
}

/* ===================== Registration / bootstrap ===================== */

/**
 * Ensure there is a sane state at (re)launch:
 * - If no state today or last was LoggedOut → append Idle now.
 */
function ensureIdleAfterLogout_() {
  const id = getCurrentAnalystId_();
  const today = toISODate_(new Date());
  const ss = master_();

  const status = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);

  const last = status[status.length - 1];
  if (last && String(last.state) !== 'LoggedOut') return false;

  // Use smart appender so previous stint (if any) is closed correctly
  appendStatusLogSmart_(id, 'Idle', 'auto', (last ? 'Relaunch after logout' : 'Day start default'));
  return true;
}

/** Ensure your Analysts row exists/updates when user starts here. */
function upsertAnalystProfile_(id, nameOpt, teamOpt) {
  const sh = getOrCreateMasterSheet_(SHEETS.ANALYSTS,
    ['analyst_id','name','team','time_zone','contracted_hours','manager']);
  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0] || []);

  let row = -1;
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) === normId_(id)) { row = r + 1; break; }
  }

  if (row === -1) {
    sh.appendRow([id, nameOpt || '', teamOpt || '', TZ, 8.5, '']);
  } else {
    if (nameOpt) sh.getRange(row, idx['name'] + 1).setValue(nameOpt);
    if (teamOpt) sh.getRange(row, idx['team'] + 1).setValue(teamOpt);
  }
}

/** Legacy register (kept for compatibility). */
function registerMe(preferredName, team) {
  const id = getCurrentAnalystId_();
  upsertAnalystProfile_(id, preferredName || '', team || '');

  const token = makeToken_();
  setSessionTokenFor_(id, token);
  PropertiesService.getUserProperties().setProperty('last_seen_iso', new Date().toISOString());

  ensureIdleAfterLogout_();
  logLoginEvent_('Login', 'Session started (web)', token);
  refreshLiveFor_(id);

  return {
    ok: true,
    analyst_id: id,
    token,
    states: STATES,
    check_types: readCheckTypes_(),
    state_info: getCurrentStateInfo_(),
    baseline_hours: getMyBaselineHours_(),
    location_today: getLocationToday_(id) || ''
  };
}

/** Takeover-friendly registration used by the UI. */
function registerMeTakeover(preferredName, team) {
  const id = getCurrentAnalystId_();
  upsertAnalystProfile_(id, preferredName||'', team||'');
  ensureUserTriggers_();
  ensureIdleAfterLogout_();

  try { dedupeLiveRowsForAnalyst_(id); } catch(e) {}

  const token = makeToken_();
  setSessionTokenFor_(id, token);
  PropertiesService.getUserProperties().setProperty('last_seen_iso', new Date().toISOString());

  logLoginEvent_('Login', 'Session started here (takeover)', token);
  refreshLiveFor_(id);

  const stateInfo = getCurrentStateInfo_();
  return {
    ok: true,
    analyst_id: id,
    token,
    states: STATES,
    check_types: readCheckTypes_(),
    state_info: stateInfo,
    baseline_hours: getMyBaselineHours_(),
    location_today: getLocationToday_(id) || ''
  };
}

/** Keep only the newest Live row for this analyst; delete older duplicates. */
function dedupeLiveRowsForAnalyst_(analystId) {
  const sh = master_().getSheetByName(SHEETS.LIVE);
  if (!sh || sh.getLastRow() < 3) return;
  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0].map(String));
  const want = normId_(analystId);

  const rows = [];
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) === want) {
      const ts = Date.parse(String(v[r][idx['last_seen_iso']] || ''));
      rows.push({ r, ts: isNaN(ts) ? 0 : ts });
    }
  }
  if (rows.length <= 1) return;

  rows.sort((a,b)=> a.ts - b.ts);
  const keep = rows[rows.length-1].r;
  for (let i = rows.length-2; i >= 0; i--) {
    sh.deleteRow(rows[i].r + 1);
  }
}

/* ===================== BACKFILL: time_in_state ===================== */
/**
 * Ensure the time_in_state column exists on StatusLogs and return its 1-based index.
 */
function ensureTimeInStateColumn_(sh) {
  const width = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1,1,1,width).getValues()[0].map(String);
  let col = header.indexOf('time_in_state');
  if (col === -1) {
    col = header.length; // next empty column (0-based)
    sh.getRange(1, col+1).setValue('time_in_state');
  }
  return col + 1; // 1-based
}

/** Enumerate YYYY-MM-DD dates inclusive. */
function enumerateDateRange_(startISO, endISO) {
  if (!startISO || !endISO) return [];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(startISO) || !/^\d{4}-\d{2}-\d{2}$/.test(endISO)) return [];
  const s = new Date(startISO + 'T00:00:00Z');
  const e = new Date(endISO + 'T00:00:00Z');
  if (e < s) return [];
  const out = [];
  for (let d = new Date(s); d <= e; d.setUTCDate(d.getUTCDate() + 1)) {
    out.push(Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'));
  }
  return out;
}

/**
 * Backfill helper: compute time_in_state per row for a date range (inclusive).
 * For each analyst/day, time_in_state = next timestamp − current timestamp.
 * If no next timestamp for that analyst the same day:
 * - if date is today → cap at now
 * - else → cap at end-of-day (23:59:59Z)
 */
/**
 * Backfill helper: compute time_in_state per row for a date range (inclusive).
 * For each analyst/day, time_in_state = next timestamp − current timestamp.
 * If no next timestamp for that analyst the same day:
 * - if date is today → cap at now
 * - else → cap at end-of-day (23:59:59Z)
 */
function backfillTimeInStateRange_(startISO, endISO) {
  // Build date list safely
  const dates = (function enumerateDateRange_(sISO, eISO){
    if (!sISO || !eISO) return [];
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sISO) || !/^\d{4}-\d{2}-\d{2}$/.test(eISO)) return [];
    const s = new Date(sISO + 'T00:00:00Z');
    const e = new Date(eISO + 'T00:00:00Z');
    if (!(s instanceof Date) || isNaN(s) || !(e instanceof Date) || isNaN(e) || e < s) return [];
    const out = [];
    for (let d = new Date(s); d <= e; d.setUTCDate(d.getUTCDate() + 1)) {
      out.push(Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'));
    }
    return out;
  })(startISO, endISO);

  if (!dates.length) return { ok:false, reason:'No dates to process' };

  // Ensure sheet & headers
  const sh = ensureStatusLogsReady_();

  // Pull all values once
  const rng = sh.getDataRange();
  const vals = rng.getValues();
  if (!vals || vals.length < 2) return { ok:true, updated:0, startISO, endISO };

  // Build header index map from the now-guaranteed headers
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(hdr);
  const tsCol = idx['timestamp_iso'];
  const dateCol = idx['date'];
  const analystCol = idx['analyst_id'];

  if (tsCol == null || dateCol == null || analystCol == null) {
    return { ok:false, reason:'StatusLogs still missing required columns after ensure', have: hdr };
  }

  // Ensure we know where to write the minutes
  let timeCol1 = idx['time_in_state'];
  if (timeCol1 == null) {
    // Shouldn’t happen because ensureStatusLogsReady_ adds it, but re-check:
    const newIdx = hdr.indexOf('time_in_state');
    timeCol1 = (newIdx === -1) ? (hdr.length + 1) : (newIdx + 1);
  } else {
    timeCol1 = timeCol1 + 1; // to A1 index
  }

  // Index rows by date for efficient processing
  const rowsByDate = {};
  for (let r = 1; r < vals.length; r++) {
    const dISO = String(vals[r][dateCol] || '');
    if (!dISO) continue;
    (rowsByDate[dISO] || (rowsByDate[dISO] = [])).push({ r, row: vals[r] });
  }

  const todayISO = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const now = new Date();
  let updates = 0;

  dates.forEach(dateISO => {
    const dayRows = rowsByDate[dateISO] || [];
    if (!Array.isArray(dayRows) || dayRows.length === 0) return;

    // Sort by analyst, then timestamp
    dayRows.sort((a,b)=>{
      const aidA = String(a.row[analystCol]||'').toLowerCase();
      const aidB = String(b.row[analystCol]||'').toLowerCase();
      if (aidA < aidB) return -1;
      if (aidA > aidB) return 1;
      const ta = Date.parse(String(a.row[tsCol]||'')) || 0;
      const tb = Date.parse(String(b.row[tsCol]||'')) || 0;
      return ta - tb;
    });

    // Compute deltas within each analyst group
    for (let i = 0; i < dayRows.length; i++) {
      const cur = dayRows[i];
      const curAid = String(cur.row[analystCol]||'').toLowerCase();
      const curTs = Date.parse(String(cur.row[tsCol]||''));
      if (!isFinite(curTs)) continue;

      // find next row for same analyst the same day
      let j = i + 1, nextTs = null;
      while (j < dayRows.length) {
        const nxt = dayRows[j];
        const nxtAid = String(nxt.row[analystCol]||'').toLowerCase();
        if (nxtAid !== curAid) break; // moved to next analyst group
        const parsed = Date.parse(String(nxt.row[tsCol]||''));
        if (isFinite(parsed)) { nextTs = parsed; break; }
        j++;
      }

      let endTs;
      if (nextTs != null) {
        endTs = nextTs;
      } else {
        // no next → cap at now if today, else EOD UTC
        endTs = (dateISO === todayISO) ? now.getTime() : Date.parse(dateISO + 'T23:59:59Z');
      }
      const mins = Math.max(0, Math.round((endTs - curTs) / 60000));

      sh.getRange(cur.r + 1, timeCol1).setValue(mins);
      updates++;
    }
  });

  return { ok:true, updated:updates, startISO, endISO };
}

/** One-click fixer: ensure headers, then backfill last 7 days. */
function Fix_StatusLogs_Headers_And_Backfill_Last7Days() {
  ensureStatusLogsReady_();
  return Backfill_TimeInState_Last7Days();
}

/** Ensure StatusLogs exists and has the required headers (adds missing ones). */
function ensureStatusLogsReady_() {
  const ss = master_();
  let sh = ss.getSheetByName(SHEETS.STATUS_LOGS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.STATUS_LOGS);
  }

  const REQUIRED = ['timestamp_iso','date','analyst_id','state','source','note','time_in_state'];

  // Read current header row (row 1). If blank, write the required headers.
  const width = Math.max(sh.getLastColumn(), REQUIRED.length);
  const row1 = sh.getRange(1, 1, 1, width).getValues()[0];
  const current = row1.map(v => String(v || '').trim());
  const isCompletelyBlank = current.every(s => s === '');

  if (isCompletelyBlank) {
    sh.getRange(1, 1, 1, REQUIRED.length).setValues([REQUIRED]);
    sh.setFrozenRows(1);
    return sh;
  }

  // Append any missing required headers to the end of row 1 (don’t remove existing).
  let changed = false;
  const next = current.slice();
  REQUIRED.forEach(h => {
    if (!next.includes(h)) { next.push(h); changed = true; }
  });
  if (changed) {
    sh.getRange(1, 1, 1, next.length).setValues([next]);
  }
  sh.setFrozenRows(1);
  return sh;
}
/* ---------- Friendly wrappers you can run from the editor ---------- */

/** Backfill a single day. */
function Backfill_TimeInState_ForDate(dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO||''))) throw new Error('Use YYYY-MM-DD');
  return backfillTimeInStateRange_(dateISO, dateISO);
}

/** Backfill an explicit inclusive range. */
function Backfill_TimeInState_Range(startISO, endISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(startISO||'')) || !/^\d{4}-\d{2}-\d{2}$/.test(String(endISO||''))) {
    throw new Error('Use YYYY-MM-DD');
  }
  return backfillTimeInStateRange_(startISO, endISO);
}

/** Backfill last 7 days (today and previous 6). */
function Backfill_TimeInState_Last7Days() {
  var tz = Session.getScriptTimeZone();
  var today = new Date();
  var endISO = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  var start = new Date(today);
  start.setDate(start.getDate() - 6);
  var startISO = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
  return backfillTimeInStateRange_(startISO, endISO);
}

/******************************************************
 * === Time-in-state backfill (robust) ===
 * Uses timestamp_iso for accuracy. Writes minutes into time_in_state.
 ******************************************************/

/** Ensure StatusLogs sheet + headers exist, return the sheet. */
function FA_ensureStatusLogsReady_() {
  const required = ['timestamp_iso','date','analyst_id','state','source','note','time_in_state'];
  const ss = master_();
  let sh = ss.getSheetByName(SHEETS.STATUS_LOGS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.STATUS_LOGS);
  }
  // align headers
  const width = Math.max(required.length, sh.getLastColumn() || 1);
  const first = sh.getRange(1, 1, 1, width).getValues()[0] || [];
  required.forEach((h, i) => { if (String(first[i] || '') !== h) sh.getRange(1, i+1).setValue(h); });
  sh.setFrozenRows(1);
  return sh;
}

/** Quick inspector (runable) – shows headers & row count. */
function Inspect_StatusLogs_Quick(){
  const sh = FA_ensureStatusLogsReady_();
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  Logger.log('StatusLogs sheet: %s | headers: %s | rows: %s',
             sh.getName(), JSON.stringify(hdr), Math.max(0, sh.getLastRow()-1));
}

/** Normalize any value to YYYY-MM-DD (TZ aware). */
function FA_toISODateSafe_(val){
  try {
    if (val instanceof Date) return toISODate_(val);
    const d = new Date(val);
    if (!isNaN(d)) return toISODate_(d);
  } catch(e){}
  return String(val || '');
}

/** Minutes between two Dates (>=0). */
function FA_minutesBetween_(a, b){
  if (!(a instanceof Date) || !(b instanceof Date)) return 0;
  const ms = b.getTime() - a.getTime();
  return ms > 0 ? Math.round(ms / 60000) : 0;
}

/**
 * Backfill time_in_state for a date range (inclusive).
 * startISO/endISO MUST be 'YYYY-MM-DD'.
 */
function Backfill_TimeInState_Range(startISO, endISO){
  if (!/^\d{4}-\d{2}-\d{2}$/.test(startISO || '') || !/^\d{4}-\d{2}-\d{2}$/.test(endISO || ''))
    throw new Error('Use YYYY-MM-DD for startISO/endISO');

  const sh = FA_ensureStatusLogsReady_();
  const v = sh.getDataRange().getValues();
  if (!v || v.length < 2) { Logger.log('No StatusLogs rows to backfill.'); return; }

  const hdr = v[0].map(String);
  const idx = indexMap_(hdr);
  ['timestamp_iso','date','analyst_id','state','time_in_state'].forEach(c=>{
    if (idx[c] == null) throw new Error('StatusLogs missing required column: '+c);
  });

  // Build lightweight rows and filter to range
  const rows = [];
  for (let r=1; r<v.length; r++){
    const tsRaw = v[r][idx['timestamp_iso']];
    const ts = new Date(tsRaw);
    if (!(ts instanceof Date) || isNaN(ts)) continue;

    const dateCell = v[r][idx['date']];
    const dateISO = FA_toISODateSafe_(dateCell || ts);
    if (dateISO < startISO || dateISO > endISO) continue;

    rows.push({
      sheetRow: r+1,
      analyst: normId_(String(v[r][idx['analyst_id']] || '')),
      dateISO,
      state: String(v[r][idx['state']] || ''),
      ts
    });
  }

  // Sort by analyst, date, timestamp
  rows.sort((a,b)=>{
    if (a.analyst !== b.analyst) return a.analyst < b.analyst ? -1 : 1;
    if (a.dateISO !== b.dateISO) return a.dateISO < b.dateISO ? -1 : 1;
    return a.ts - b.ts;
  });

  const colTimeInState = idx['time_in_state'] + 1;
  const todayISO = toISODate_(new Date());
  const now = new Date();

  let updated = 0, scanned = rows.length;

  // Walk rows and compute end boundary = next row (same analyst+date) or EOD/now
  for (let i=0; i<rows.length; i++){
    const cur = rows[i];
    // find next row within same analyst+date
    let j = i+1;
    let nextTs = null;
    if (j < rows.length && rows[j].analyst === cur.analyst && rows[j].dateISO === cur.dateISO){
      nextTs = rows[j].ts;
    }

    if (!nextTs){
      // end of group: to end of day (23:59) or now if today
      const endBoundary = (cur.dateISO === todayISO)
        ? now
        : new Date(cur.dateISO + 'T23:59:59');
      nextTs = endBoundary;
    }

    const mins = FA_minutesBetween_(cur.ts, nextTs);
    sh.getRange(cur.sheetRow, colTimeInState).setValue(mins);
    updated++;
  }

  Logger.log('Backfill complete. Rows scanned: %s; In range: %s; Updated: %s; Wrote column index: %s',
             v.length-1, scanned, updated, colTimeInState);
}

/** Convenience wrappers you can run from the UI. */
function Backfill_TimeInState_Last7Days(){
  const endISO = toISODate_(new Date());
  const startDate = new Date(endISO + 'T00:00:00');
  startDate.setDate(startDate.getDate() - 6); // last 7 days inclusive
  const startISO = toISODate_(startDate);
  Backfill_TimeInState_Range(startISO, endISO);
}

function Backfill_TimeInState_ForDate(dateISO){
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateISO || '')) throw new Error('Use YYYY-MM-DD');
  Backfill_TimeInState_Range(dateISO, dateISO);
}

