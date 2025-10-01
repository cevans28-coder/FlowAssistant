/******************************************************
 * 10_ui_api.gs — UI endpoints for Analyst app (slim)
 *
 * What lives here now:
 * - Client config (getClientConfig)
 * - App init payload (getInitData)
 * - Today summary for the sidebar (getTodaySummary)
 * - Update baseline hours (updateMyBaselineHours)
 * - Check types list (readCheckTypes_)
 * - Baseline hours getter (getMyBaselineHours_)
 * - Diagnostics for Exceptions (optional)
 *
 * What does NOT live here anymore (to avoid duplication):
 * - setState() → use 04_state_engine.gs
 * - _bustTodaySummaryCache_ → 04_state_engine.gs
 * - WorkLocations V2 helpers (getTodayWorkPlan_, setMyLocationToday, etc.)
 * → use your Work Plan v2 file (getWorkingLocationsV2Sheet_, getTodayWorkPlan_, saveMyWorkPlan, getMyWorkPlan, getLocationToday_)
 * - _hm_toMinutes_ → defined in Work Plan v2 file
 *
 * Depends on:
 * - TZ, SHEETS, STATES (00_constants.gs)
 * - toISODate_, normId_, indexMap_, minutesBetween_, readRows_ (01/02 utils)
 * - master_() (02_master_access.gs)
 * - getCurrentAnalystId_(), makeToken_(), setSessionTokenFor_(),
 * requireSession_(), heartbeat(), ensureIdleAfterLogout_() (03_sessions.gs)
 * - getCurrentStateInfo_(), refreshLiveFor_() (04_state_engine.gs)
 * - getAcceptedMeetingMinutes_() (06_calendar.gs)
 * - getTodayWorkPlan_(), getLocationToday_() (Work Plan v2 file)
 ******************************************************/

/* ========================= Settings / Config ========================= */

/** Read simple client config from MASTER → Settings (key | value). */
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

/* ============================ App bootstrap ============================ */

/**
 * Initial payload for the web UI / sidebar.
 * - Ensures a sane starting state (Idle if first launch / after LoggedOut)
 * - Creates/adopts a fresh session token (and heartbeats)
 * - Returns static lists + current state + baseline + today’s saved location/plan
 */
function getInitData() {
  const id = getCurrentAnalystId_();

  // Baseline guard for first-run / post-logout sessions
  try { ensureIdleAfterLogout_(); } catch (e) {}

  // Fresh session token + heartbeat (so Live gets online/last_seen updated)
  const token = makeToken_();
  setSessionTokenFor_(id, token);
  try { heartbeat(token); } catch (e) {}

  // Work plan for today comes from Work Plan v2
  const plan = getTodayWorkPlan_(id);

  return {
    analyst_id: id,
    token,
    states: STATES,
    check_types: readCheckTypes_(),
    state_info: getCurrentStateInfo_(),
    baseline_hours: getMyBaselineHours_(),
    location_today: (plan && plan.location) || getLocationToday_(id) || '',
    work_plan_today: plan || null
  };
}

/* ============================ Today summary ============================ */

/**
 * Sum exception minutes for an analyst on a given date (robust to email/local-part + Date cells).
 */
function getExceptionMinutesForAnalystDate_(analystId, dateISO){
  if (!analystId || !/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO||''))) return 0;

  const ss = master_();
  const tz = Session.getScriptTimeZone();

  // locate Exceptions sheet
  const names = [];
  if (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.EXCEPTIONS) names.push(SHEETS.EXCEPTIONS);
  names.push('Exceptions');
  let sh = null;
  for (const n of names){ const s = ss.getSheetByName(n); if (s && s.getLastRow()>1){ sh = s; break; } }
  if (!sh) return 0;

  const vals = sh.getDataRange().getValues();
  const hdr = (vals[0]||[]).map(String);
  const H = {}; hdr.forEach((h,i)=> H[String(h).trim().toLowerCase()] = i);
  const col = name => (H[name] != null ? H[name] : -1);

  const cAid = col('analyst_id') >= 0 ? col('analyst_id') : col('analyst_email');
  const cDate = col('date_iso');
  const cStart = col('start_ts');
  const cEnd = col('end_ts');
  const cMin = col('minutes');
  const cStat = col('status');

  if (cAid < 0) return 0;

  const normPerson = s => {
    const raw = String(s||'').trim().toLowerCase();
    if (!raw) return '';
    const m = raw.match(/^([^@]+)@/);
    return (m ? m[1] : raw);
  };
  const want = normPerson(analystId);

  const toISOday = v => {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const s = String(v||'').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const d = new Date(s);
    return isNaN(d) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  };

  const num = v => {
    if (v == null || v === '') return 0;
    const s = String(v).replace(',','.');
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  };

  let total = 0;

  for (let r=1; r<vals.length; r++){
    const row = vals[r];

    if (normPerson(row[cAid]) !== want) continue;

    const st = cStat >= 0 ? String(row[cStat]||'').toUpperCase() : 'NEW';
    if (st === 'REJECTED' || st === 'CANCELED') continue;

    let rowDay = '';
    if (cDate >= 0 && row[cDate]) rowDay = toISOday(row[cDate]);
    if (!rowDay && cStart >= 0 && row[cStart]) rowDay = toISOday(row[cStart]);
    if (rowDay !== dateISO) continue;

    let mins = 0;
    if (cMin >= 0 && row[cMin] !== '' && row[cMin] != null){
      mins = num(row[cMin]);
    } else if (cStart >= 0 && cEnd >= 0 && row[cStart] && row[cEnd]){
      const s = (row[cStart] instanceof Date) ? row[cStart] : new Date(row[cStart]);
      const e = (row[cEnd] instanceof Date) ? row[cEnd] : new Date(row[cEnd]);
      if (e > s) mins = Math.round((e - s)/60000);
    }

    if (mins > 0) total += mins;
  }

  return Math.max(0, Math.round(total));
}

/**
 * Today summary for the signed-in analyst (read-only).
 * Returns:
 * - meeting_mins: accepted events from CAL_PULL
 * - working_mins_calc: planned_mins − meeting_mins (never < 0)
 * - output_total: count of CheckEvents for today
 * - state/since: latest StatusLogs entry for today
 * - location_today + work_plan: from WorkLocations_v2
 * - exception_mins: minutes of Exceptions today
 */
function getTodaySummary(){
  const id = getCurrentAnalystId_();
  const today = toISODate_(new Date());
  const key = 'SUM:' + id + ':' + today;

  const cache = CacheService.getUserCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  try {
    const ss = master_();

    // Checks today
    const checks = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
      .filter(r => r.date_str === today && r.analyst_id_norm === id);

    // Meetings (accepted)
    const meetingMins = Math.max(0, Number(getAcceptedMeetingMinutes_(id, today) || 0));

    // Work plan (start/end/lunch) -> planned minutes for today
    const plan = getTodayWorkPlan_(id);
    const plannedMins = (plan && typeof plan.planned_mins === 'number')
      ? Math.max(0, Math.round(plan.planned_mins))
      : Math.round(7.5 * 60); // default day = 7.5h

    // Exceptions (sum of today's minutes)
    const exceptionMins = Math.max(0, Number(getExceptionMinutesForAnalystDate_(id, today) || 0));

    // Base working minutes (PLAN - meetings)
    const workingMinsCalc = Math.max(0, plannedMins - meetingMins);

    // Current state + since
    const statusRows = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
      .filter(r => r.date_str === today && r.analyst_id_norm === id)
      .sort((a,b) => a.ts - b.ts);
    const last = statusRows[statusRows.length - 1] || null;

    const out = {
      ok: true,
      date: today,
      output_total: checks.length,
      meeting_mins: meetingMins,
      working_mins_calc: workingMinsCalc,
      exception_mins: exceptionMins,
      state: last ? String(last.state || 'Idle') : 'Idle',
      since_iso: (last && last.ts) ? last.ts.toISOString() : null,
      location_today: (plan && plan.location) || getLocationToday_(id) || '',
      planned_work_mins: plannedMins, // for display/debug
      work_plan: plan || null // for display/debug
    };

    try { cache.put(key, JSON.stringify(out), 45); } catch(e){}
    return out;

  } catch (e) {
    const fallback = {
      ok: false,
      date: today,
      output_total: 0,
      meeting_mins: 0,
      working_mins_calc: Math.round(7.5 * 60),
      exception_mins: 0,
      state: 'Idle',
      since_iso: null,
      error: String((e && e.message) || e)
    };
    try { cache.put(key, JSON.stringify(fallback), 30); } catch(e2){}
    return fallback;
  }
}

/* ========================= Baseline hours (Analysts) ========================= */

function updateMyBaselineHours(hours) {
  const hrs = Number(hours);
  if (!hrs || hrs <= 0) throw new Error('Please enter a positive number of hours.');

  const sh = master_().getSheetByName(SHEETS.ANALYSTS);
  if (!sh) throw new Error('Analysts sheet missing.');

  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0] || {});
  const id = getCurrentAnalystId_();
  const idCol = idx['analyst_id'], hrsCol = idx['contracted_hours'];

  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idCol]) === id) {
      if (hrsCol != null) sh.getRange(r + 1, hrsCol + 1).setValue(hrs);
      refreshLiveFor_(id);
      return { ok: true, hours: hrs };
    }
  }

  // Append if not found
  sh.appendRow([id, '', '', TZ, hrs, '']);
  refreshLiveFor_(id);
  return { ok: true, hours: hrs };
}

function readCheckTypes_() {
  const sh = master_().getSheetByName(SHEETS.CHECK_TYPES);
  if (!sh || sh.getLastRow() <= 1) return [];
  const v = sh.getDataRange().getValues();

  const norm = s => String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  return v.slice(1)
    .filter(r => r[0])
    .map(r => ({
      name: String(r[0]),
      name_norm: norm(r[0]),
      avg_minutes: Number(r[1]) || 0
    }));
}

function getMyBaselineHours_() {
  const sh = master_().getSheetByName(SHEETS.ANALYSTS);
  if (!sh || sh.getLastRow() <= 1) return 7.5;

  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0] || {});
  const id = getCurrentAnalystId_();

  const idCol = idx['analyst_id'] ?? idx['email'] ?? idx['id'];
  const hCol =
    (idx['contracted_hours'] != null ? idx['contracted_hours'] :
    (idx['baseline_hours'] != null ? idx['baseline_hours'] :
    (idx['hours_per_day'] != null ? idx['hours_per_day'] : null)));

  if (idCol == null || hCol == null) return 7.5;

  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idCol]) === id) {
      const h = Number(v[r][hCol]);
      return h > 0 ? h : 7.5;
    }
  }
  return 7.5;
}

/* ====================== Diagnostics (optional) ====================== */

function DIAG_ListTodayExceptionsForMe() {
  const id = getCurrentAnalystId_();
  const day = toISODate_(new Date());
  const ss = master_();

  const names = (typeof SHEETS !== 'undefined' && SHEETS.EXCEPTIONS) ? [SHEETS.EXCEPTIONS] : [];
  names.push('Exceptions');

  let sh = null;
  for (const n of names) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) { sh = s; break; }
  }
  if (!sh) { Logger.log('No Exceptions sheet found'); return; }

  const v = sh.getDataRange().getValues();
  const hdr = v[0].map(String);
  const idx = {};
  hdr.forEach((h,i)=> idx[String(h).trim().toLowerCase()] = i);

  function col(name){ return idx[name] ?? -1; }

  const cAid = col('analyst_id') >= 0 ? col('analyst_id') : col('analyst_email');
  const cDate = col('date_iso');
  const cStart = col('start_ts');
  const cEnd = col('end_ts');
  const cMin = col('minutes');
  const cStatus= col('status');

  const want = normId_(id);
  let total = 0;

  Logger.log('--- Exceptions for %s on %s ---', id, day);

  for (let r=1;r<v.length;r++){
    const row = v[r];
    const aidRaw = String(row[cAid] || '');
    const aidNorm = normId_(aidRaw.includes('@') ? aidRaw : aidRaw);
    if (aidNorm !== want) continue;

    const st = cStatus >=0 ? String(row[cStatus]||'').toUpperCase() : 'NEW';
    if (st === 'REJECTED' || st === 'CANCELED') {
      Logger.log('Skip row %s: status=%s', r+1, st);
      continue;
    }

    let d = '';
    if (cDate >= 0 && row[cDate]) {
      d = String(row[cDate]).slice(0,10);
    } else if (cStart >= 0 && row[cStart]) {
      const dt = row[cStart] instanceof Date ? row[cStart] : new Date(row[cStart]);
      if (!isNaN(dt)) d = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (d !== day) continue;

    let mins = 0, reason = '';
    if (cMin >= 0 && row[cMin] !== '' && row[cMin] != null) {
      mins = Number(row[cMin]) || 0;
      reason = 'minutes column';
    } else if (cStart >=0 && cEnd >=0 && row[cStart] && row[cEnd]) {
      const s = row[cStart] instanceof Date ? row[cStart] : new Date(row[cStart]);
      const e = row[cEnd] instanceof Date ? row[cEnd] : new Date(row[cEnd]);
      if (e > s) { mins = Math.round((e - s)/60000); reason = 'derived from start/end'; }
      else { reason = 'invalid start/end'; }
    } else {
      reason = 'no minutes and no start/end';
    }

    Logger.log('Row %s | status=%s | date=%s | mins=%s (%s)', r+1, st, d, mins, reason);
    total += mins;
  }

  Logger.log('TOTAL minutes today (diag): %s', total);
}
