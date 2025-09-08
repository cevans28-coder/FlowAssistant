/******************************************************
 * 07_metrics.gs — Daily metrics compute & writers (optimized)
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_, indexMap_, minutesBetween_, readRows_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getAcceptedMeetingMinutes_() (06_calendar.gs)
 * - readCheckTypes_(), getBaselineHoursForAnalyst_() (10_ui_api / 00_constants)
 * - upsertLive_() (04_state_engine.gs)
 * - upsertDailyTypeSummaryFor_, rebuildDailyTypeSummaryForDateForAll_ (14_rollups.gs)
 *
 * Notes:
 * - DailyMetrics caps utilisation at 100% (as agreed).
 * - Live KPIs are computed with “available minutes = baseline − meetings”.
 ******************************************************/

/* ======================== SHEET GUARDIAN ======================== */
/**
 * Ensure and return the canonical DailyMetrics sheet with fixed headers.
 * If the tab exists but headers drifted, we realign row 1 safely.
 */
function getDailyMetricsSheet_() {
  const ss = master_();
  const HEADERS = [
    'date', // yyyy-MM-dd (TZ)
    'analyst_id', // normalized email
    'available_mins', // baseline − meeting mins (>=0)
    'handling_mins', // sum(CheckEvents.duration_mins)
    'output_total', // count of checks
    'standard_mins', // Σ (count_by_type × CheckTypes.avg_minutes)
    'efficiency_pct', // (handling / standard) × 100
    'utilisation_pct', // (handling / available) × 100 (capped at 100)
    'throughput_per_hr', // output_total / (available_mins / 60)
    'flags', // reserved for QA notes/flags
    'notes', // free notes
    'meeting_mins' // accepted meeting minutes from CalendarPull
  ];

  // Exact tab first
  let sh = ss.getSheetByName(SHEETS.DAILY);
  if (!sh) {
    // Fallback: adopt any tab starting with "DailyMetrics"
    const candidate = ss.getSheets().find(s => s.getName().trim().toLowerCase().startsWith('dailymetrics'));
    sh = candidate || ss.insertSheet(SHEETS.DAILY);
    sh.setName(SHEETS.DAILY);
  }

  // Ensure header row is present & aligned
  const width = Math.max(HEADERS.length, sh.getLastColumn() || 1);
  const first = sh.getRange(1, 1, 1, width).getValues()[0] || [];
  HEADERS.forEach((h, i) => {
    if (String(first[i] || '').trim() !== h) sh.getRange(1, i + 1).setValue(h);
  });
  sh.setFrozenRows(1);
  return sh;
}

/* =================== CORE COMPUTATION (ONE ROW) =================== */
/**
 * Compute one metrics row OBJECT for a given analyst & date.
 * Returns a POJO whose keys match getDailyMetricsSheet_() headers.
 */
function computeDailyMetricsRow_(analystId, dateISO) {
  const ss = master_();

  // 1) Baseline hours → minutes (fallback to analyst baseline or 8.5h)
  const baselineHours =
    Number(getBaselineHoursForAnalyst_(analystId)) ||
    Number(getMyBaselineHours_ && getMyBaselineHours_()) || 8.5;
  const baselineMins = Math.max(0, Math.round(baselineHours * 60));

  // 2) Meetings (accepted) for the date → minutes
  const meetingMins = Math.max(0, Number(getAcceptedMeetingMinutes_(analystId, dateISO) || 0));

  // 3) Available minutes = baseline − meetings (>=0)
  const availableMins = Math.max(0, baselineMins - meetingMins);

  // 4) CheckEvents → handling + counts by type (single pass)
  const ce = ss.getSheetByName(SHEETS.CHECK_EVENTS);
  let handlingMins = 0;
  let outputTotal = 0;
  const perTypeCount = {};
  if (ce && ce.getLastRow() > 1) {
    const vals = ce.getDataRange().getValues();
    const idx = indexMap_(vals[0].map(String));
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      if (String(row[idx['date']] || '') !== dateISO) continue;
      if (normId_(row[idx['analyst_id']] || '') !== normId_(analystId)) continue;

      const mins = Number(row[idx['duration_mins']] || 0);
      if (mins > 0) handlingMins += mins;

      const ct = String(row[idx['check_type']] || '');
      if (ct) {
        perTypeCount[ct] = (perTypeCount[ct] || 0) + 1;
        outputTotal += 1;
      }
    }
  }

  // 5) Standard mins = Σ (count_by_type × avg_minutes from CheckTypes)
  const types = readCheckTypes_(); // [{name, avg_minutes}]
  const avgMap = {};
  (types || []).forEach(t => (avgMap[String(t.name)] = Number(t.avg_minutes || 0)));
  let standardMins = 0;
  Object.keys(perTypeCount).forEach(ct => {
    standardMins += (perTypeCount[ct] || 0) * (avgMap[ct] || 0);
  });

  // 6) KPIs (guard divisions; utilisation capped at 100 for DailyMetrics)
  const efficiencyPct = (handlingMins > 0 && standardMins > 0)
    ? Math.round((handlingMins / standardMins) * 100)
    : 0;

  const utilisationRaw = availableMins > 0 ? (handlingMins / availableMins) * 100 : 0;
  const utilisationPct = Math.min(100, Math.round(utilisationRaw)); // agreed cap

  const throughputPerHr = (availableMins > 0)
    ? Number((outputTotal / (availableMins / 60)).toFixed(2))
    : 0;

  return {
    date: dateISO,
    analyst_id: analystId,
    available_mins: Math.max(0, Math.round(availableMins)),
    handling_mins: Math.max(0, Math.round(handlingMins)),
    output_total: outputTotal,
    standard_mins: Math.max(0, Math.round(standardMins)),
    efficiency_pct: efficiencyPct,
    utilisation_pct: utilisationPct,
    throughput_per_hr: throughputPerHr,
    flags: '',
    notes: '',
    meeting_mins: Math.max(0, Math.round(meetingMins))
  };
}

/* ====================== WRITER / APPEND HELPERS ====================== */
/**
 * Append a single metrics row OBJECT to DailyMetrics using header mapping.
 * Unknown keys are ignored; missing keys remain blank.
 */
function appendDailyMetrics_(rowObj) {
  const sh = getDailyMetricsSheet_();
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(hdr);
  const out = new Array(hdr.length).fill('');

  Object.keys(rowObj || {}).forEach(k => {
    if (Object.prototype.hasOwnProperty.call(idx, k)) out[idx[k]] = rowObj[k];
  });

  sh.appendRow(out);
}

/* ============================ PUBLIC APIS ============================ */
/** Build today's metrics for *me* (menu/UI button) + update per-type rollup. */
function buildMyMetricsToday() {
  const analystId = getCurrentAnalystId_();
  const dateISO = toISODate_(new Date());
  const row = computeDailyMetricsRow_(analystId, dateISO);
  appendDailyMetrics_(row);

  // NEW: keep per-type summary in sync
  try { upsertDailyTypeSummaryFor_(analystId, dateISO); } catch(e) {}

  return { ok: true, analyst_id: analystId, date: dateISO };
}

/** Build metrics for *me* for an explicit date (UI button with token) + rollup. */
function buildMyMetricsForDate(token, dateISO) {
  requireSession_(token);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) throw new Error('Use YYYY-MM-DD');
  const analystId = getCurrentAnalystId_();
  const row = computeDailyMetricsRow_(analystId, dateISO);
  appendDailyMetrics_(row);

  // NEW: keep per-type summary in sync
  try { upsertDailyTypeSummaryFor_(analystId, dateISO); } catch(e) {}

  return { ok: true, analyst_id: analystId, date: dateISO };
}

/** Build metrics for ALL analysts for a given date (admin/automation) + rollup. */
function rebuildMetricsForDateForAll(dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) throw new Error('Use YYYY-MM-DD');

  const ss = master_();
  const aSh = ss.getSheetByName(SHEETS.ANALYSTS);
  if (!aSh || aSh.getLastRow() < 2) throw new Error('No analysts found.');

  const vals = aSh.getDataRange().getValues();
  const idx = indexMap_(vals[0].map(String));

  for (let r = 1; r < vals.length; r++) {
    const aid = String(vals[r][idx['analyst_id']] || '').trim();
    if (!aid) continue;
    const rowObj = computeDailyMetricsRow_(aid, dateISO);
    appendDailyMetrics_(rowObj);

    // NEW: keep per-type summary in sync
    try { upsertDailyTypeSummaryFor_(aid, dateISO); } catch(e) {}
  }
  return { ok: true, date: dateISO };
}

/** Build metrics for a specific analyst & date (TL tooling) + rollup. */
function buildMetricsForAnalystDate(analystId, dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) throw new Error('Use YYYY-MM-DD');
  if (!analystId) throw new Error('Missing analystId');
  const rowObj = computeDailyMetricsRow_(analystId, dateISO);
  appendDailyMetrics_(rowObj);

  // NEW: keep per-type summary in sync
  try { upsertDailyTypeSummaryFor_(analystId, dateISO); } catch(e) {}

  return { ok: true, analyst_id: analystId, date: dateISO };
}

/* ====================== LIVE KPI COMPUTATION ====================== */
function _faSafeNum(n){ n = Number(n); return Number.isFinite(n) ? n : 0; }
function _faClamp0(n){ n = _faSafeNum(n); return n < 0 ? 0 : n; }
function _faNormKey(s){
  return String(s || '')
    .replace(/\u00A0/g, ' ') // NBSP → space
    .replace(/\s+/g, ' ') // collapse spaces
    .trim()
    .toLowerCase();
}

/**
 * TODAY snapshot for Live.
 * Efficiency (%) = (standard / handling) × 100 (LIVE flavour)
 * Utilisation (%) = (handling / available) × 100
 * TPH = output / (available/60)
 */
function computeLiveProductionToday_(analystId) {
  const ss = master_();
  const id = normId_(analystId);
  const dateISO = toISODate_(new Date());

  // Baseline & meetings → available
  const baselineHours = _faSafeNum(getBaselineHoursForAnalyst_(id) || 8.5);
  const baselineMins = _faClamp0(Math.round(baselineHours * 60));
  const meetingMins = _faClamp0(getAcceptedMeetingMinutes_(id, dateISO));
  const availableMins = _faClamp0(baselineMins - meetingMins);

  // Logged-in minutes today (everything except LoggedOut/OOO), cap open stint at now
  const sl = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
              .filter(r => r.date_str === dateISO && r.analyst_id_norm === id)
              .sort((a,b)=> a.ts - b.ts);
  const now = new Date();
  let loggedInMins = 0;
  for (let i=0;i<sl.length;i++){
    const cur = sl[i], nxt = sl[i+1];
    if (!cur.ts) continue;
    let s = cur.ts, e = nxt && nxt.ts ? nxt.ts : now;
    if (e > now) e = now;
    if (!(s instanceof Date) || isNaN(s) || !(e instanceof Date) || isNaN(e) || e <= s) continue;
    const st = String(cur.state||'');
    if (!/^(LoggedOut|OOO)$/i.test(st)) loggedInMins += _faClamp0(minutesBetween_(s, e));
  }
  loggedInMins = Math.round(loggedInMins);

  // Checks today → handling mins & per-type counts (normalised key)
  const ceRows = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
                  .filter(r => r.date_str === dateISO && r.analyst_id_norm === id);
  let handlingMins = 0;
  const perType = {};
  ceRows.forEach(r => {
    handlingMins += _faClamp0(r.duration_mins);
    const key = _faNormKey(r.check_type);
    if (key) perType[key] = (perType[key] || 0) + 1;
  });
  handlingMins = Math.round(handlingMins);
  const outputTotal = ceRows.length;

  // CheckTypes averages (normalised keys)
  const avgMap = {};
  (readCheckTypes_() || []).forEach(t => {
    const key = _faNormKey(t.name_norm || t.name);
    if (key) avgMap[key] = _faClamp0(t.avg_minutes);
  });

  let standardMins = 0;
  Object.keys(perType).forEach(key => {
    standardMins += _faClamp0(perType[key]) * _faClamp0(avgMap[key]);
  });
  standardMins = Math.round(standardMins);

  // KPIs (LIVE flavour)
  const efficiency = (handlingMins > 0 && standardMins > 0)
    ? Math.round((standardMins / handlingMins) * 100)
    : 0;

  const utilisation = (availableMins > 0)
    ? Math.round((handlingMins / availableMins) * 100)
    : 0;

  const tph = (availableMins > 0)
    ? Number((outputTotal / (availableMins / 60)).toFixed(2))
    : 0;

  return {
    dateISO,
    logged_in_mins: _faClamp0(loggedInMins),
    handling_mins: _faClamp0(handlingMins),
    standard_mins: _faClamp0(standardMins),
    available_mins: _faClamp0(availableMins),
    live_efficiency_pct: _faClamp0(efficiency),
    live_utilisation_pct: _faClamp0(utilisation),
    live_throughput_per_hr: _faSafeNum(tph)
  };
}

/**
 * Compute "logged-in" minutes for an analyst on a given date.
 * Logged-in = all state time except LoggedOut / OOO.
 * Uses the existing timeline to avoid re-implementing stint logic.
 */
function computeLoggedInMinutesForDay_(analystId, dateISO) {
  if (!analystId || !/^\d{4}-\d{2}-\d{2}$/.test(dateISO)) return 0;
  // If local timeline is exposed:
  var tl = getAnalystRangeTimeline(analystId, dateISO, dateISO);
  // If only via library, swap to:
  // var tl = QATracker.getAnalystRangeTimeline(analystId, dateISO, dateISO);

  if (!tl || !tl.days || !tl.days.length) return 0;
  var day = tl.days[0];
  var mins = 0;
  (day.stints || []).forEach(function(s) {
    var st = String(s.state || '').toLowerCase();
    if (st === 'loggedout' || st === 'ooo') return;
    var start = new Date(s.start_iso).getTime();
    var end = new Date(s.end_iso).getTime();
    if (isFinite(start) && isFinite(end) && end > start) {
      mins += Math.round((end - start) / 60000);
    }
  });
  return Math.max(0, mins|0);
}

/**
 * Single source of truth: compute today’s LIVE KPIs → write to Live sheet.
 * Idempotent and safe to call from heartbeat, check logging, and TL actions.
 */
function updateLiveKPIsFor_(analystId) {
  const id = normId_(analystId || getCurrentAnalystId_());
  const k = computeLiveProductionToday_(id);

  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.LIVE);
  if (!sh || sh.getLastRow() < 2) return;

  const v = sh.getDataRange().getValues();
  const idx = indexMap_(v[0].map(String));

  // Find row for analyst
  let rowIndex = -1;
  for (let r = 1; r < v.length; r++) {
    if (normId_(v[r][idx['analyst_id']]) === id) { rowIndex = r + 1; break; }
  }
  if (rowIndex === -1) return;

  // Safe numeric sets
  const safeSet = (colName, value) => {
    if (idx[colName] != null) sh.getRange(rowIndex, idx[colName] + 1).setValue(Number(value) || 0);
  };

  safeSet('logged_in_mins', k.logged_in_mins);
  safeSet('live_efficiency_pct', k.live_efficiency_pct);
  safeSet('live_utilisation_pct', k.live_utilisation_pct);
  safeSet('live_throughput_per_hr', k.live_throughput_per_hr);
}

/* ====================== NIGHTLY SAFETY NET (STEP 3) ====================== */
/**
 * Rebuild DailyTypeSummary for YESTERDAY for all analysts.
 * Create a time-driven trigger to run this nightly as a safety net.
 */
function cron_RebuildTypeSummary_Yesterday(){
  const tz = Session.getScriptTimeZone();
  const y = Utilities.formatDate(new Date(Date.now() - 24*60*60*1000), tz, 'yyyy-MM-dd');
  try {
    rebuildDailyTypeSummaryForDateForAll_(y);
  } catch (e) {
    Logger.log('cron_RebuildTypeSummary_Yesterday error: ' + (e && e.message));
  }
}

/**
 * One-off helper to create the nightly trigger at ~01:10 local time.
 * Run this once from the Apps Script editor.
 */
function createNightlyTypeSummaryTrigger_(){
  // Delete any existing duplicates for cleanliness (optional)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'cron_RebuildTypeSummary_Yesterday')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('cron_RebuildTypeSummary_Yesterday')
    .timeBased()
    .everyDays(1)
    .atHour(1) // 01:10 local time
    .nearMinute(10)
    .create();

  Logger.log('Nightly trigger created for cron_RebuildTypeSummary_Yesterday at ~01:10.');
}
