/******************************************************
 * 03_sessions.gs — Sessions, heartbeats, smart StatusLogs (v2)
 *
 * Exposes:
 * - getMyEmail_(), getCurrentAnalystId_(), makeToken_()
 * - setSessionTokenFor_(analystId, token), getSessionTokenFor_(analystId)
 * - requireSession_(token), heartbeat(token)
 * - ensureIdleAfterLogout_()
 * - appendStatusLogSmart_(analystId, state, source, note, ts)
 * - reconcileTimeInStateForToday_(analystId)
 * - logLoginEvent_(analystId, event, note, token)
 * - registerMeTakeover(emailOpt, nameOpt) // safe no-op registration helper
 *
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - normId_, toISODate_, minutesBetween_, indexMap_, readRows_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - upsertLive_(), refreshLiveFor_() (04_state_engine.gs)
 ******************************************************/

/* ------------------- tiny local helpers ------------------- */
function getMyEmail_(){
  try {
    return (Session.getActiveUser().getEmail() ||
            Session.getEffectiveUser().getEmail() || '').trim();
  } catch(e){ return ''; }
}

function getCurrentAnalystId_() {
  try {
    const em =
      (Session.getActiveUser().getEmail() ||
       Session.getEffectiveUser().getEmail() || '')
      .trim().toLowerCase();
    return em || '';
  } catch (e) {
    return '';
  }
}
function makeToken_(){
  return Utilities.getUuid();
}

function _getStatusLogsSheet_(){
  // Canonical v2 sheet; fall back to any sheet that matches "StatusLogs"
  const ss = master_();
  let sh = ss.getSheetByName(SHEETS.STATUS_LOGS);
  if (sh) return sh;

  const cand = ss.getSheets().find(s => String(s.getName()).toLowerCase().indexOf('statuslogs') === 0);
  if (cand) return cand;

  // create canonical with v2 headers if missing
  return getOrCreateMasterSheet_(SHEETS.STATUS_LOGS, [
    'timestamp_iso','date','analyst_id','state','source','note','time_in_state'
  ]);
}

function _getLoginHistorySheet_(){
  return getOrCreateMasterSheet_(SHEETS.LOGIN_HISTORY || 'LoginHistory', [
    'timestamp_iso','date','analyst_id','event','note','session_token'
  ]);
}

/* ------------------- session token storage ------------------- */
/** Set/replace the session token for an analyst (persists to Live). */
function setSessionTokenFor_(analystId, token){
  const id = normId_(analystId || getCurrentAnalystId_());
  const tok = String(token || '');
  // Write token + touch last_seen to “now”
  const nowIso = new Date().toISOString();
  upsertLive_(id, { session_token: tok, last_seen_iso: nowIso });
  try {
    // Keep 04_state_engine.online computation happy
    PropertiesService.getUserProperties().setProperty('last_seen_iso', nowIso);
  } catch(e){}
  return tok;
}

/** Fetch the current session token for an analyst from Live ('' if none). */
function getSessionTokenFor_(analystId){
  const id = normId_(analystId || getCurrentAnalystId_());
  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.LIVE);
  if (!sh || sh.getLastRow() < 2) return '';
  const vals = sh.getDataRange().getValues();
  const idx = indexMap_(vals[0].map(String));
  for (let r=1; r<vals.length; r++){
    if (normId_(vals[r][idx['analyst_id']] || '') === id){
      return String(vals[r][idx['session_token']] || '');
    }
  }
  return '';
}

/** Throws if token is missing/does not match the current user’s Live token. */
function requireSession_(token){
  const id = getCurrentAnalystId_();
  if (!id) throw new Error('No signed-in user.');
  const liveTok = getSessionTokenFor_(id);
  if (!token || !liveTok || String(token) !== String(liveTok)){
    throw new Error('Your session has expired. Please reopen the add-on.');
  }
}

/**
 * requireSessionOrAdopt_(token)
 * - Validates the session token.
 * - If it's expired/missing, auto-issues a fresh token for the current user,
 * heartbeats it, and returns the new token string.
 * - Returns the validated (or newly adopted) token.
 */
function requireSessionOrAdopt_(token) {
  try {
    requireSession_(token);
    return String(token || '');
  } catch (e) {
    var msg = (e && e.message || '').toLowerCase();
    // Heuristics for "expired / missing / invalid token"
    if (msg.includes('expired') || msg.includes('token') || msg.includes('session')) {
      var id = getCurrentAnalystId_();
      var fresh = makeToken_();
      setSessionTokenFor_(id, fresh);
      try { heartbeat(fresh); } catch (_) {}
      return fresh;
    }
    throw e;
  }
}

/* ------------------- heartbeats & login events ------------------- */
/**
 * Heartbeat from UI. Updates last_seen, keeps Live fresh, and writes a LoginHistory “Heartbeat”.
 * Note: state is not changed here; presence is derived from last_seen + state in 04_state_engine.
 */
function heartbeat(token){
  requireSession_(token);

  const id = getCurrentAnalystId_();
  const now = new Date();
  const nowIso = now.toISOString();

  try { PropertiesService.getUserProperties().setProperty('last_seen_iso', nowIso); } catch(e){}

  // Touch Live (just last_seen; 04.refreshLiveFor_ will compute online)
  upsertLive_(id, { last_seen_iso: nowIso });

  // Optionally refresh live snapshot & KPIs (best-effort)
  try { refreshLiveFor_(id); } catch(e){}
  try { updateLiveKPIsFor_(id); } catch(e){}

  // Audit
  logLoginEvent_(id, 'Heartbeat', '', getSessionTokenFor_(id));

  return { ok:true, ts: nowIso };
}

/** Append an audit row to LoginHistory. */
function logLoginEvent_(analystId, event, note, token){
  const sh = _getLoginHistorySheet_();
  const now = new Date();
  sh.appendRow([now.toISOString(), toISODate_(now), String(analystId||''), String(event||''), String(note||''), String(token||'')]);
}

/* ------------------- status logs (smart writer) ------------------- */
/**
 * Append a new state row and close the previous stint (sets time_in_state in minutes).
 * Safe to call repeatedly; only the *previous* row’s time_in_state is computed here.
 *
 * @param {string} analystId
 * @param {string} state
 * @param {string} source e.g. 'UI', 'system', 'manager:<email>', 'watchdog'
 * @param {string} note
 * @param {Date=} tsOpt timestamp for the new row (defaults to now)
 */
function appendStatusLogSmart_(analystId, state, source, note, tsOpt){
  const id = normId_(analystId || getCurrentAnalystId_());
  const ts = tsOpt instanceof Date ? tsOpt : new Date();
  const dateISO = toISODate_(ts);

  const sh = _getStatusLogsSheet_();

  // Load today’s rows for this analyst (minimal read)
  const rows = readRows_(sh)
    .filter(r => r.date_str === dateISO && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);

  // If there is a previous row, set its time_in_state to delta to “ts”
  if (rows.length){
    const prev = rows[rows.length - 1];
    // Find its 1-based sheet row index
    const hdr = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
    const idx = indexMap_(hdr);
    // We can locate the exact row via a scan; cost is trivial for a day slice
    const all = sh.getDataRange().getValues();
    for (let r = 1; r < all.length; r++){
      const dateStr = String(all[r][idx['date']] || '');
      const aidStr = normId_(all[r][idx['analyst_id']] || '');
      const tsStr = String(all[r][idx['timestamp_iso']] || '');
      if (dateStr === dateISO && aidStr === id && tsStr && prev.ts && tsStr === prev.ts.toISOString()){
        // compute minutes
        const mins = Math.max(0, minutesBetween_(prev.ts, ts));
        const col = idx['time_in_state'];
        if (col != null) sh.getRange(r+1, col + 1).setValue(mins);
        break;
      }
    }
  }

  // Append the new row with empty time_in_state (to be closed by next change or backfill)
  sh.appendRow([ts.toISOString(), dateISO, id, String(state||''), String(source||''), String(note||''), '']);

  // Keep Live coherent (state + since)
  try {
    upsertLive_(id, {
      state: String(state||''),
      since_iso: ts.toISOString(),
      last_seen_iso: new Date().toISOString()
    });
  } catch(e){}

  // Audit (optional)
  try { logLoginEvent_(id, 'StateChange', (source||'') + (note?(' — ' + note):''), getSessionTokenFor_(id)); } catch(e){}
}

/**
 * Recompute *today’s* time_in_state values for an analyst.
 * Fills all but the last row (left open); if today, the last stint ends “now”.
 */
function reconcileTimeInStateForToday_(analystId){
  const id = normId_(analystId || getCurrentAnalystId_());
  const sh = _getStatusLogsSheet_();
  const today = toISODate_(new Date());

  const rows = readRows_(sh)
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a,b)=> a.ts - b.ts);

  if (!rows.length) return { ok:true, updated:0 };

  const hdr = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  const idx = indexMap_(hdr);
  let updated = 0;

  for (let i=0; i<rows.length; i++){
    const cur = rows[i];
    const next = rows[i+1];
    if (!cur.ts) continue;

    // Find row index in sheet
    // (A small linear scan is fine for day-level reconciliation)
    const all = sh.getDataRange().getValues();
    for (let r=1; r<all.length; r++){
      const dateStr = String(all[r][idx['date']] || '');
      const aidStr = normId_(all[r][idx['analyst_id']] || '');
      const tsStr = String(all[r][idx['timestamp_iso']] || '');
      if (dateStr === today && aidStr === id && tsStr && cur.ts && tsStr === cur.ts.toISOString()){
        const end = next && next.ts ? next.ts : new Date();
        const mins = Math.max(0, minutesBetween_(cur.ts, end));
        if (idx['time_in_state'] != null){
          sh.getRange(r+1, idx['time_in_state'] + 1).setValue(mins);
          updated++;
        }
        break;
      }
    }
  }
  return { ok:true, updated };
}

/* ------------------- launch helper ------------------- */
/**
 * Ensure a safe starting state for today (used in getInitData()).
 * - If no StatusLogs for today → write Idle
 * - If last state for today is LoggedOut → write Idle
 */
function ensureIdleAfterLogout_(){
  const id = getCurrentAnalystId_();
  if (!id) return;

  const sh = _getStatusLogsSheet_();
  const today = toISODate_(new Date());
  const rows = readRows_(sh)
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a,b)=> a.ts - b.ts);

  const last = rows[rows.length - 1];
  if (!last || String(last.state||'') === 'LoggedOut'){
    appendStatusLogSmart_(id, 'Idle', 'system', 'init baseline', new Date());
  }
}

/* ------------------- lightweight registration (optional) ------------------- */
/**
 * Safe registration helper (used in SMOKE_Heartbeat).
 * Creates/refreshes your Live row and gives you a token if you don’t have one yet.
 */
function registerMeTakeover(emailOpt, nameOpt){
  const id = getCurrentAnalystId_();
  if (!id) return { ok:false, reason:'no user' };

  // ensure a Live row exists (name/team are best-effort from arguments)
  try { upsertLive_(id, { name: String(nameOpt||''), last_seen_iso: new Date().toISOString() }); } catch(e){}

  // create a token if missing
  let tok = getSessionTokenFor_(id);
  if (!tok){
    tok = makeToken_();
    setSessionTokenFor_(id, tok);
  }

  // Write a login event
  logLoginEvent_(id, 'Register', 'registerMeTakeover', tok);
  return { ok:true, analyst_id:id, token: tok };
}
