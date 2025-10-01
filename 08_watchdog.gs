/******************************************************
 * 08_watchdog.gs — Inactivity rules & gentle nudges (optimized)
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, minutesBetween_, normId_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), getSessionTokenFor_(), logLoginEvent_() (03_sessions.gs)
 * - upsertLive_() (04_state_engine.gs)
 *
 * Triggers (installed by ensureUserTriggers_ in 99_triggers.gs):
 * - watchdog_heartbeat_10min() — every ~10 minutes per user
 *
 * Behaviour:
 * - If no heartbeat in >10m while Working/Admin → set Idle
 * - If Idle and >60m without heartbeat → LoggedOut
 * - If Break/Lunch and >60m without heartbeat → LoggedOut
 * - All other states (Meeting/Training/Coaching/OOO/…) with >10m gap → LoggedOut
 *
 * Notes:
 * - Writes ONE StatusLogs row only when a transition is required.
 * - Keeps Live (v1) in sync and audits in LoginHistory.
 * - Optimized: reads only the needed columns from StatusLogs.
 ******************************************************/

/** Fast lookup of today's latest state row for an analyst (reads 4 columns only). */
function _getLatestStateTodayFast_(ss, dateISO, analystIdNorm) {
  const sh = ss.getSheetByName(SHEETS.STATUS_LOGS);
  if (!sh || sh.getLastRow() < 2) return { state: 'Idle', ts: null, source: '' };

  // Header map
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const H = (name)=> hdr.indexOf(name);
  const cDate = H('date'), cAid = H('analyst_id'), cTs = H('timestamp_iso'), cState = H('state'), cSrc = H('source');
  if (cDate < 0 || cAid < 0 || cTs < 0 || cState < 0) return { state: 'Idle', ts: null, source: '' };

  const nRows = sh.getLastRow() - 1;
  // Pull just these columns as arrays
  const colDate = sh.getRange(2, cDate + 1, nRows, 1).getValues().map(r => String(r[0] || ''));
  const colAid = sh.getRange(2, cAid + 1, nRows, 1).getValues().map(r => normId_(r[0] || ''));
  const colTs = sh.getRange(2, cTs + 1, nRows, 1).getValues().map(r => new Date(r[0]));
  const colState = sh.getRange(2, cState + 1, nRows, 1).getValues().map(r => String(r[0] || ''));
  const colSrc = (cSrc >= 0)
    ? sh.getRange(2, cSrc + 1, nRows, 1).getValues().map(r => String(r[0] || ''))
    : new Array(nRows).fill('');

  let bestIdx = -1;
  let bestTime = 0;
  for (let i = 0; i < nRows; i++) {
    if (colDate[i] !== dateISO) continue;
    if (colAid[i] !== analystIdNorm) continue;
    const t = colTs[i] instanceof Date && !isNaN(colTs[i]) ? colTs[i].getTime() : 0;
    if (t >= bestTime) { bestTime = t; bestIdx = i; }
  }

  if (bestIdx === -1) return { state: 'Idle', ts: null, source: '' };
  return {
    state: colState[bestIdx] || 'Idle',
    ts: (colTs[bestIdx] instanceof Date && !isNaN(colTs[bestIdx])) ? colTs[bestIdx] : null,
    source: colSrc[bestIdx] || ''
  };
}

function watchdog_heartbeat_10min() {
  // Only enforce Idle > 60m => LoggedOut. Do not infer idle from missing heartbeats.
  var ss = master_();
  var id = getCurrentAnalystId_();

  var live = ss.getSheetByName(SHEETS.LIVE);
  if (!live || live.getLastRow() < 2) return;

  var v = live.getDataRange().getValues();
  var idx = indexMap_(v[0].map(String));
  var row = null;
  for (var r=1; r<v.length; r++){
    if (normId_(v[r][idx['analyst_id']]) === id){ row = v[r]; break; }
  }
  if (!row) return;

  var lastState = String(row[idx['state']] || 'Idle');
  var sinceIso = String(row[idx['since_iso']] || '');
  if (lastState !== 'Idle' || !sinceIso) return;

  var since = new Date(sinceIso);
  if (!(since instanceof Date) || isNaN(since)) return;

  var idleMins = minutesBetween_(since, new Date());
  if (idleMins >= 60) {
    appendStatusLogSmart_(id, 'LoggedOut', 'watchdog', 'Auto-logout: Idle ≥ 60m', new Date());
    try { refreshLiveFor_(id); } catch(e){}
    try { updateLiveKPIsFor_(id); } catch(e){}
  }
}

/**
 * Nudge: if you’re in an accepted calendar meeting right now but your state is Working,
 * send a gentle email reminder. Uses CalendarPull_v2 via SHEETS.CAL_PULL.
 */
function nudgeIfInMeetingButWorking_() {
  try {
    const now = new Date();
    const dateISO = toISODate_(now);
    const id = getCurrentAnalystId_();
    const ss = master_();

    // CalendarPull_v2 (SHEETS.CAL_PULL should point to the v2 tab)
    const sh = ss.getSheetByName(SHEETS.CAL_PULL);
    if (!sh || sh.getLastRow() < 2) return;

    const v = sh.getDataRange().getValues();
    const idx = (function(h){ const m={}; h.forEach((x,i)=>m[String(x)] = i); return m; })(v[0].map(String));
    const cDate = idx['date'], cAid = idx['analyst_id'], cStart = idx['start_iso'], cEnd = idx['end_iso'], cStatus = idx['my_status'];
    if ([cDate,cAid,cStart,cEnd,cStatus].some(x=> x==null)) return;

    const idNorm = normId_(id);
    let inMeeting = false;
    for (let r=1; r<v.length; r++){
      if (String(v[r][cDate]) !== dateISO) continue;
      if (normId_(v[r][cAid]) !== idNorm) continue;
      const status = String(v[r][cStatus] || '').toLowerCase();
      if (!['accepted','accept','yes','y'].includes(status)) continue;
      const s = new Date(v[r][cStart]); const e = new Date(v[r][cEnd]);
      if (s instanceof Date && !isNaN(s) && e instanceof Date && !isNaN(e) && now >= s && now <= e) { inMeeting = true; break; }
    }
    if (!inMeeting) return;

    // Latest state fast
    const latest = _getLatestStateTodayFast_(ss, dateISO, idNorm);
    if (String(latest.state || '') !== 'Working') return;

    const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    if (!email) return;
    MailApp.sendEmail({
      to: email,
      subject: 'Horizon — quick reminder',
      htmlBody: 'We detected you appear to be in a meeting, but your state is <b>Working</b>. Want to switch to <b>Meeting</b>?'
    });
  } catch (e) {
    // swallow to keep trigger healthy
  }
}
