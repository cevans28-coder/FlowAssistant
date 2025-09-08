/******************************************************
 * 10_ui_api.gs — UI endpoints for Analyst app
 *
 * What lives here:
 * - Client config (getClientConfig)
 * - App init payload (getInitData)
 * - Today summary for the sidebar (getTodaySummary)
 * - Update baseline hours (updateMyBaselineHours)
 * - Check types list (readCheckTypes_)
 * - Baseline hours getter (getMyBaselineHours_)
 * - Working location get/set (getLocationToday_, setMyLocationToday)
 *
 * Depends on (from your project):
 * - TZ, SHEETS, STATES (00_constants.gs)
 * - toISODate_, normId_, indexMap_, minutesBetween_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), readRows_() (02_master_access.gs)
 * - getCurrentAnalystId_(), makeToken_(), setSessionTokenFor_(),
 * requireSession_(), heartbeat(), ensureIdleAfterLogout_() (03_sessions.gs)
 * - refreshLiveFor_() (04_state_engine.gs)
 * - getAcceptedMeetingMinutes_() (06_calendar.gs)
 ******************************************************/

/**
 * Read simple client config from MASTER → Settings.
 * Row format: key | value
 * - WEB_APP_URL | https://script.google.com/.../exec
 */
function getClientConfig() {
  const sh = master_().getSheetByName(SHEETS.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return { web_app_url: '' };

  const v = sh.getDataRange().getValues();
  for (let r = 1; r < v.length; r++) {
    const key = String(v[r][0] || '').trim().toUpperCase();
    if (key === 'WEB_APP_URL') return { web_app_url: String(v[r][1] || '').trim() };
  }
  return { web_app_url: '' };
}

/**
 * Initial payload for the web UI / sidebar.
 * - Ensures a sane starting state for today (Idle if first launch / after LoggedOut)
 * - Creates/adopts a fresh session token
 * - Sends back static lists + current state + baseline + today’s saved location
 */
function getInitData() {
  const id = getCurrentAnalystId_();

  // 1) Ensure we have a baseline state for today (idempotent)
  try { ensureIdleAfterLogout_(); } catch (e) {}

  // 2) Issue a new token and heartbeat; adopt if previous is stale
  const token = makeToken_();
  setSessionTokenFor_(id, token);
  try { heartbeat(token); } catch (e) {}

  // 3) Current state info and preferences
  return {
    analyst_id: id,
    token,
    states: STATES,
    check_types: readCheckTypes_(),
    state_info: getCurrentStateInfo_(), // from state engine
    baseline_hours: getMyBaselineHours_(), // from Analysts
    location_today: getLocationToday_(id) || ''
  };
}

/**
 * Today summary for the signed-in analyst (read-only, no side-effects).
 * Returns:
 * - meeting_mins: accepted events from CAL_PULL
 * - working_mins_calc: baseline*60 − meeting_mins (never < 0)
 * - output_total: count of CheckEvents for today
 * - state/since: latest StatusLogs entry for today
 * - location_today: saved once per day
 */
function getTodaySummary(){
  const id = getCurrentAnalystId_();
  const key = 'SUM:' + id + ':' + toISODate_(new Date());
  const cache = CacheService.getUserCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  // your existing try/catch body:
  try{
    const ss = master_();
    const dateISO = toISODate_(new Date());
    const checks = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS)).filter(r=> r.date_str===dateISO && r.analyst_id_norm===id);
    const cal = readRows_(ss.getSheetByName(SHEETS.CAL_PULL)).filter(r=> r.date_str===dateISO && r.analyst_id_norm===id && String(r.my_status)==='YES');

    let meetingMins = 0; cal.forEach(r=>{ if (r.start && r.end) meetingMins += minutesBetween_(r.start,r.end); });
    const workingMinsCalc = Math.max(0, Math.round(getMyBaselineHours_()*60 - meetingMins));

    const status = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS)).filter(r=> r.date_str===dateISO && r.analyst_id_norm===id).sort((a,b)=> a.ts-b.ts);
    const last = status[status.length-1];
    const location_today = getLocationToday_(id) || '';

    const out = {
      ok:true, date:dateISO, output_total:checks.length, meeting_mins:meetingMins,
      working_mins_calc:workingMinsCalc, state: last?String(last.state):'Idle',
      since_iso: last&&last.ts? last.ts.toISOString(): null, location_today
    };
    try { cache.put(key, JSON.stringify(out), 45); } catch(e){}
    return out;
  }catch(e){
    const out = { ok:false, date:toISODate_(new Date()), output_total:0, meeting_mins:0,
      working_mins_calc:Math.round(getMyBaselineHours_()*60), state:'Idle', since_iso:null, error:String(e && e.message || e) };
    try { cache.put(key, JSON.stringify(out), 30); } catch(e2){}
    return out;
  }
}

/**
 * Update my baseline (contracted) hours in MASTER → Analysts.
 * Also refreshes Live so TL/UI views pick up the change.
 */
function updateMyBaselineHours(hours) {
  const hrs = Number(hours);
  if (!hrs || hrs <= 0) throw new Error('Please enter a positive number of hours.');

  const sh = master_().getSheetByName(SHEETS.ANALYSTS);
  if (!sh) throw new Error('Analysts sheet missing.');

  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0] || {});
  const id = getCurrentAnalystId_();
  const idCol = idx['analyst_id'], hrsCol = idx['contracted_hours'];

  // Update if exists
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idCol]) === id) {
      if (hrsCol != null) sh.getRange(r + 1, hrsCol + 1).setValue(hrs);
      refreshLiveFor_(id);
      return { ok: true, hours: hrs };
    }
  }

  // Else append a new row with sane defaults
  const row = [id, '', '', TZ, hrs, ''];
  sh.appendRow(row);
  refreshLiveFor_(id);
  return { ok: true, hours: hrs };
}

/**
 * Read CheckTypes for the UI (dropdown + averages if needed later).
 * Sheet columns: check_type | avg_minutes | sla_minutes | weight
 * Returns: [{ name, avg_minutes }]
 */
function readCheckTypes_() {
  const sh = master_().getSheetByName(SHEETS.CHECK_TYPES);
  if (!sh || sh.getLastRow() <= 1) return [];
  const v = sh.getDataRange().getValues();

  // Normalise helper: lower-case, trim, collapse spaces, strip NBSP
  const norm = s => String(s || '')
    .replace(/\u00A0/g, ' ') // NBSP → normal space
    .replace(/\s+/g, ' ') // collapse runs of space
    .trim()
    .toLowerCase();

  return v.slice(1)
    .filter(r => r[0])
    .map(r => ({
      name: String(r[0]),
      name_norm: norm(r[0]), // <— add a precomputed normalised key
      avg_minutes: Number(r[1]) || 0
    }));
}

/**
 * Return my baseline hours quickly (defaults to 8.5 if not set).
 */
function getMyBaselineHours_() {
  const sh = master_().getSheetByName(SHEETS.ANALYSTS);
  if (!sh || sh.getLastRow() <= 1) return 8.5;

  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0] || {});
  const id = getCurrentAnalystId_();

  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) === id) {
      const h = Number(v[r][idx['contracted_hours']]);
      return h > 0 ? h : 8.5;
    }
  }
  return 8.5;
}

/* ========== Working location (Home/Office), once per day ========== */

/**
 * Get today’s saved location for a given analyst ('' if not set).
 */
function getLocationToday_(analystId) {
  const sh = master_().getSheetByName(SHEETS.LOCATION);
  if (!sh || sh.getLastRow() <= 1) return '';

  const rows = readRows_(sh)
    .filter(r => r.date_str === toISODate_(new Date()) &&
                 r.analyst_id_norm === normId_(analystId));

  if (!rows.length) return '';
  const last = rows[rows.length - 1];
  return String(last.location || last['location']) || '';
}

/** Convenience for UI bound to the current user. */
function getLocationToday_analyst() {
  return getLocationToday_(getCurrentAnalystId_());
}

/**
 * Set my working location (Home/Office) once per day.
 * - Enforces one set per day; throws if already set
 * - Appends to WorkLocations and refreshes Live
 */
function setMyLocationToday(location) {
  const loc = String(location || '').trim();
  if (!['Home', 'Office'].includes(loc)) throw new Error('Location must be Home or Office.');

  const id = getCurrentAnalystId_();
  const today = toISODate_(new Date());
  const sh = getOrCreateMasterSheet_(SHEETS.LOCATION,
    ['timestamp_iso','date','analyst_id','location','source','note']);

  // If already set, block a second write
  const already = readRows_(sh).some(r =>
    r.analyst_id_norm === normId_(id) && r.date_str === today
  );
  if (already) throw new Error('Working location already set for today.');

  sh.appendRow([new Date().toISOString(), today, id, loc, 'UI', '']);
  refreshLiveFor_(id);
  return { ok: true, location: loc };
}
