/******************************************************
 * 03_sessions.gs — Identity & session lifecycle
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - normId_, toISODate_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), indexMap_() (02_master_access.gs)
 * - upsertLive_(), refreshLiveFor_() (04_state_engine.gs) // referenced by heartbeat/logOff
 ******************************************************/

/** Resolve the current user's analyst_id (normalised email). */
function getCurrentAnalystId_() {
  const active = Session.getActiveUser() && Session.getActiveUser().getEmail();
  if (active && active.indexOf('@') > -1) return normId_(active);

  const eff = Session.getEffectiveUser() && Session.getEffectiveUser().getEmail();
  if (eff && eff.indexOf('@') > -1) return normId_(eff);

  const saved = PropertiesService.getUserProperties().getProperty('analyst_id');
  return normId_(saved || 'unknown_user');
}

/** Convenience: active or effective email (raw, not normalised). */
function getMyEmail_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
}

/** Generate a new random session token. */
function makeToken_() {
  return Utilities.getUuid();
}

/** Read the current session token for an analyst from Live. */
function getSessionTokenFor_(id) {
  const sh = master_().getSheetByName(SHEETS.LIVE);
  if (!sh) return '';
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return '';
  const idx = indexMap_(v[0]);
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) === id) {
      return String(v[r][idx['session_token']] || '');
    }
  }
  return '';
}

/** Write/update session token for an analyst in Live (creates row if missing). */
function setSessionTokenFor_(id, token) {
  const sh = master_().getSheetByName(SHEETS.LIVE);
  if (!sh) return;
  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0] || {});
  let row = -1;

  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) === id) { row = r + 1; break; }
  }

  const col = (idx['session_token'] || 10) + 1; // safe default if headers present
  if (row === -1) {
    // Create minimal Live row, then write token
    upsertLive_(id, { analyst_id: id });
    const v2 = sh.getDataRange().getValues();
    const idx2 = indexMap_(v2[0] || {});
    const col2 = (idx2['session_token'] || 10) + 1;
    const row2 = sh.getLastRow();
    sh.getRange(row2, col2).setValue(token);
  } else {
    sh.getRange(row, col).setValue(token);
  }
}

/**
 * Clear a user's session (token, online flag, state to LoggedOut).
 * Called by logOff() and watchdog when force-logging out.
 */
function clearSession_(id) {
  const sh = master_().getSheetByName(SHEETS.LIVE);
  if (!sh) return;
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return;

  const idx = indexMap_(v[0]);
  let row = -1;
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) === id) { row = r + 1; break; }
  }
  const nowIso = new Date().toISOString();
  if (row !== -1) {
    if (idx['session_token'] !== undefined) sh.getRange(row, idx['session_token'] + 1).setValue('');
    if (idx['online'] !== undefined) sh.getRange(row, idx['online'] + 1).setValue('NO');
    if (idx['last_seen_iso'] !== undefined) sh.getRange(row, idx['last_seen_iso'] + 1).setValue(nowIso);
    if (idx['state'] !== undefined) sh.getRange(row, idx['state'] + 1).setValue('LoggedOut');
    if (idx['since_iso'] !== undefined) sh.getRange(row, idx['since_iso'] + 1).setValue(nowIso);
  }
  PropertiesService.getUserProperties().deleteProperty('last_seen_iso');
}

/**
 * Require a valid session token for the current user.
 * If no token is passed, we return ok (used for some background calls).
 * If a prior token exists but appears stale/offline, we adopt the new one.
 * Throws if another live session is active elsewhere.
 */
function requireSession_(token) {
  if (!token) return { ok: true, analyst_id: getCurrentAnalystId_(), no_token: true };

  const id = getCurrentAnalystId_();
  const current = getSessionTokenFor_(id);
  if (!current || token === current) return { ok: true, analyst_id: id };

  // Check Live row recency to see if we can adopt the new token safely
  const sh = master_().getSheetByName(SHEETS.LIVE);
  if (sh) {
    const v = sh.getDataRange().getValues();
    if (v.length > 1) {
      const idx = indexMap_(v[0]);
      for (let r = 1; r < v.length; r++) {
        if (normId_(v[r][idx['analyst_id']]) === id) {
          const liveToken = String(v[r][idx['session_token']] || '');
          const liveState = String(v[r][idx['state']] || '');
          const lastSeen = new Date(String(v[r][idx['last_seen_iso']] || new Date(0).toISOString()));
          const offline = minutesBetween_(lastSeen, new Date()) > 5;

          if (!liveToken || liveState === 'LoggedOut' || offline) {
            setSessionTokenFor_(id, token); // adopt the new token
            return { ok: true, analyst_id: id };
          }
          throw new Error('You are already logged in elsewhere. Please log off on the other device first.');
        }
      }
    }
  }

  // Fallback: if we couldn't find a row, adopt
  setSessionTokenFor_(id, token);
  return { ok: true, analyst_id: id };
}

/**
 * Heartbeat — marks user as recently active and nudges Live.
 * UI calls this every ~60–90s.
 */
function heartbeat(token) {
  requireSession_(token);
  PropertiesService.getUserProperties().setProperty('last_seen_iso', new Date().toISOString());
  refreshLiveFor_(getCurrentAnalystId_()); // comes from 04_state_engine.gs
  try {
  const id = getCurrentAnalystId_();
  const k = computeLiveProductionToday_(id);
  upsertLive_(id, {
    logged_in_mins: k.logged_in_mins,
    live_efficiency_pct: k.live_efficiency_pct,
    live_utilisation_pct: k.live_utilisation_pct,
    live_throughput_per_hr: k.live_throughput_per_hr
  });
} catch(e) {}

  return { ok: true };
  try { updateLiveKPIsFor_(id); } catch(e) {}
}

/**
 * User-initiated logout from the UI.
 * Writes an audited StatusLogs row, clears session, refreshes Live.
 */
function logOff(token, note) {
  requireSession_(token);
  const id = getCurrentAnalystId_();
  const ts = new Date();

  const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS,
    ['timestamp_iso','date','analyst_id','state','source','note']);
  sh.appendRow([ts.toISOString(), toISODate_(ts), id, 'LoggedOut', 'UI', note || 'User log off']);

  logLoginEvent_('Logout', note || 'User log off', token);
  clearSession_(id);
  refreshLiveFor_(id); // reflect in Live immediately
  return { ok: true };
}

/**
 * Append a LoginHistory row for audit (session events, watchdog, etc.).
 */
function logLoginEvent_(event, note, token) {
  const sh = getOrCreateMasterSheet_(SHEETS.LOGIN_HISTORY,
    ['timestamp_iso','date','analyst_id','event','note','session_token']);
  const id = getCurrentAnalystId_();
  const ts = new Date();
  sh.appendRow([ts.toISOString(), toISODate_(ts), id, event, note || '', token || getSessionTokenFor_(id)]);
}

/**
 * First-time setup on registration:
 * - ensures analyst exists in Analysts
 * - installs per-user triggers (moved to 99_triggers later)
 * - creates session + heartbeat + login history
 */
/**
 * Register/start a session.
 * If the analyst is NEW, preferredName and team are REQUIRED.
 * If the analyst exists, we only update name/team when they are provided.
 */
function registerMe(analystIdOptional, preferredNameOptional, teamOptional) {
  const sh = getOrCreateMasterSheet_(SHEETS.ANALYSTS, ['analyst_id','name','team','time_zone','contracted_hours','manager']);
  const id = normId_(analystIdOptional || getCurrentAnalystId_());
  const preferredName = String(preferredNameOptional || '').trim();
  const team = String(teamOptional || '').trim();

  const vals = sh.getDataRange().getValues();
  const idx = indexMap_(vals[0] || []);
  let foundRow = -1;

  for (let r = 1; r < vals.length; r++) {
    if (normId_(vals[r][idx['analyst_id']]) === id) { foundRow = r + 1; break; }
  }

  if (foundRow === -1) {
    // New analyst → require both fields
    if (!preferredName) throw new Error('Please enter your Preferred Name to complete registration.');
    if (!team) throw new Error('Please enter your Team to complete registration.');
    sh.appendRow([id, preferredName, team, TZ, 8.5, '']);
  } else {
    // Existing → upsert if provided (don’t overwrite with blanks)
    if (preferredName) sh.getRange(foundRow, idx['name'] + 1).setValue(preferredName);
    if (team) sh.getRange(foundRow, idx['team'] + 1).setValue(team);
    // Ensure timezone/hours defaults exist
    if (!vals[foundRow - 1][idx['time_zone']]) sh.getRange(foundRow, idx['time_zone'] + 1).setValue(TZ);
    if (!Number(vals[foundRow - 1][idx['contracted_hours']])) sh.getRange(foundRow, idx['contracted_hours'] + 1).setValue(8.5);
  }

  // Normal session bootstrapping
  ensureUserTriggers_();
  ensureIdleAfterLogout_();

  const token = makeToken_();
  setSessionTokenFor_(id, token);
  heartbeat(token);
  logLoginEvent_('Login','Session started', token);

  // Return name/team so UI can show it immediately
  return { ok:true, analyst_id:id, token, name: preferredName || '', team: team || '' };
}

/**
 * On (re)launch, if no state for today or last state is LoggedOut,
 * insert an Idle row so timers/UI have a baseline. (Idempotent.)
 */
function ensureIdleAfterLogout_() {
  const id = getCurrentAnalystId_();
  const today = toISODate_(new Date());
  const ss = master_();

  const status = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);

  const last = status[status.length - 1];
  const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS,
    ['timestamp_iso','date','analyst_id','state','source','note']);

  if (!last) {
    sh.appendRow([new Date().toISOString(), today, id, 'Idle', 'auto', 'Day start default']);
    return true;
  }
  if (String(last.state) === 'LoggedOut') {
    sh.appendRow([new Date().toISOString(), today, id, 'Idle', 'auto', 'Relaunch after logout']);
    return true;
  }
  return false;
}

// Ensure your Analysts row exists/updates when user starts here
function upsertAnalystProfile_(id, nameOpt, teamOpt) {
  const sh = getOrCreateMasterSheet_(SHEETS.ANALYSTS, ['analyst_id','name','team','time_zone','contracted_hours','manager']);
  const v = sh.getDataRange().getValues(); const idx = indexMap_(v[0]||[]);
  let row = -1;
  for (let r=1;r<v.length;r++) if (normId_(v[r][idx['analyst_id']])===normId_(id)) { row=r+1; break; }
  if (row === -1) {
    sh.appendRow([id, nameOpt||'', teamOpt||'', TZ, 8.5, '']);
  } else {
    if (nameOpt) sh.getRange(row, idx['name']+1).setValue(nameOpt);
    if (teamOpt) sh.getRange(row, idx['team']+1).setValue(teamOpt);
  }
}

// NEW: always claim the session *here* (overwrites any other token)
function registerMeTakeover(preferredName, team) {
  const id = getCurrentAnalystId_();
  upsertAnalystProfile_(id, preferredName||'', team||'');
  ensureUserTriggers_(); // your existing helper
  ensureIdleAfterLogout_(); // your existing helper

  const token = makeToken_();
  setSessionTokenFor_(id, token); // write THIS token to Live
  PropertiesService.getUserProperties().setProperty('last_seen_iso', new Date().toISOString());

  logLoginEvent_('Login', 'Session started here (takeover)', token); // audit trail
  refreshLiveFor_(id); // reflect Live row

  // Return same shape as your existing init
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
