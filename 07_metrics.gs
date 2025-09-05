/******************************************************
 * 07_metrics.gs — Daily metrics compute & writers
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_, indexMap_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getAcceptedMeetingMinutes_() (06_calendar.gs)
 * - readCheckTypes_(), getBaselineHoursForAnalyst_() // defined elsewhere in your project
 ******************************************************/

/**
 * Ensure and return the canonical DailyMetrics sheet with your exact headers.
 * Headers:
 * date, analyst_id, available_mins, handling_mins, output_total,
 * standard_mins, efficiency_pct, utilisation_pct, throughput_per_hr, flags, notes
 */
function getDailyMetricsSheet_() {
  const ss = master_();
  const HEADERS = [
    'date',
    'analyst_id',
    'available_mins',
    'handling_mins',
    'output_total',
    'standard_mins',
    'efficiency_pct',
    'utilisation_pct',
    'throughput_per_hr',
    'flags',
    'notes'
  ];

  // Exact tab name first
  let sh = ss.getSheetByName(SHEETS.DAILY);
  if (sh) {
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sh.setFrozenRows(1);
    }
    return sh;
  }

  // Fallback: adopt any "DailyMetrics..." and rename
  const candidate = ss.getSheets().find(s => s.getName().trim().toLowerCase().startsWith('dailymetrics'));
  if (candidate) {
    candidate.setName(SHEETS.DAILY);
    if (candidate.getLastRow() === 0) {
      candidate.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      candidate.setFrozenRows(1);
    }
    return candidate;
  }

  // Create once if missing
  sh = ss.insertSheet(SHEETS.DAILY);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Compute one metrics row OBJECT for a given analyst & date.
 * Returns fields matching the headers in getDailyMetricsSheet_().
 */
function computeDailyMetricsRow_(analystId, dateISO) {
  const ss = master_();

  // 1) Available minutes = baseline - accepted meetings
  const baselineHours = Number(getBaselineHoursForAnalyst_(analystId) || 8.5);
  const baselineMins = Math.round(baselineHours * 60);
  const meetingMins = getAcceptedMeetingMinutes_(analystId, dateISO);
  const availableMins = Math.max(0, baselineMins - meetingMins);

  // 2) From CheckEvents: handling_mins, output_total
  const ce = ss.getSheetByName(SHEETS.CHECK_EVENTS);
  let handlingMins = 0;
  let outputTotal = 0;
  const perTypeCount = {}; // e.g., { "CDD Check": 2, ... }

  if (ce && ce.getLastRow() > 1) {
    const vals = ce.getDataRange().getValues();
    const hdr = vals[0].map(String);
    const idx = indexMap_(hdr);

    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      if (String(row[idx['date']]) !== dateISO) continue;
      if (normId_(row[idx['analyst_id']]) !== normId_(analystId)) continue;

      const mins = Number(row[idx['duration_mins']] || 0);
      if (mins > 0) handlingMins += mins;

      const ct = String(row[idx['check_type']] || '');
      if (ct) {
        perTypeCount[ct] = (perTypeCount[ct] || 0) + 1;
        outputTotal += 1;
      }
    }
  }

  // 3) Standard minutes = Σ (count × avg_minutes) from CheckTypes
  const types = readCheckTypes_(); // [{ name, avg_minutes }, ...]
  const avgMap = {};
  types.forEach(t => { avgMap[String(t.name)] = Number(t.avg_minutes || 0); });

  let standardMins = 0;
  Object.keys(perTypeCount).forEach(ct => {
    standardMins += (perTypeCount[ct] || 0) * (avgMap[ct] || 0);
  });

  // 4) KPIs (guard divisions)
  const efficiencyPct = standardMins > 0 ? Math.round((handlingMins / standardMins) * 100) : 0;
  const utilisationPct = availableMins > 0 ? Math.round((handlingMins / availableMins) * 100) : 0;
  const throughputPerHr = availableMins > 0 ? Number((outputTotal / (availableMins / 60)).toFixed(2)) : 0;

  return {
    date: dateISO,
    analyst_id: analystId,
    available_mins: availableMins,
    handling_mins: handlingMins,
    output_total: outputTotal,
    standard_mins: standardMins,
    efficiency_pct: efficiencyPct,
    utilisation_pct: utilisationPct,
    throughput_per_hr: throughputPerHr,
    flags: '',
    notes: ''
  };
}

/**
 * Append a single metrics row OBJECT to DailyMetrics.
 */
function appendDailyMetrics_(rowObj) {
  const dm = getDailyMetricsSheet_();
  dm.appendRow([
    rowObj.date,
    rowObj.analyst_id,
    rowObj.available_mins,
    rowObj.handling_mins,
    rowObj.output_total,
    rowObj.standard_mins,
    rowObj.efficiency_pct,
    rowObj.utilisation_pct,
    rowObj.throughput_per_hr,
    rowObj.flags || '',
    rowObj.notes || ''
  ]);
}

/**
 * Public: Build today's metrics for the signed-in analyst (menu/UI button).
 */
function buildMyMetricsToday() {
  const analystId = getCurrentAnalystId_();
  const dateISO = toISODate_(new Date());
  const row = computeDailyMetricsRow_(analystId, dateISO);
  appendDailyMetrics_(row);
  return { ok: true, analyst_id: analystId, date: dateISO };
}

/**
 * Public: Build metrics for a picked date for the signed-in analyst (UI button).
 * Requires a session token (to align with your UI flow).
 */
function buildMyMetricsForDate(token, dateISO) {
  requireSession_(token);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateISO)) throw new Error('Use YYYY-MM-DD');

  const analystId = getCurrentAnalystId_();
  const row = computeDailyMetricsRow_(analystId, dateISO);
  appendDailyMetrics_(row);
  return { ok: true, analyst_id: analystId, date: dateISO };
}

/**
 * Admin: Build one metrics row for a given analyst & date (used by TL actions).
 */
function buildMetricsForAnalystDate(analystId, dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateISO)) throw new Error('Use YYYY-MM-DD');
  if (!analystId) throw new Error('Missing analystId');

  const row = computeDailyMetricsRow_(analystId, dateISO);
  appendDailyMetrics_(row);
  return { ok: true, analyst_id: analystId, date: dateISO };
}

/**
 * Admin: Rebuild metrics for ALL analysts for the given date.
 * Appends one row per analyst (does not overwrite).
 */
function rebuildMetricsForDateForAll(dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateISO)) throw new Error('Use YYYY-MM-DD');

  const ss = master_();
  const dm = getDailyMetricsSheet_();
  const aSh = ss.getSheetByName(SHEETS.ANALYSTS);
  if (!aSh) throw new Error('Analysts sheet missing.');

  const aVals = aSh.getDataRange().getValues();
  if (aVals.length <= 1) throw new Error('No analysts found.');
  const aHdr = aVals[0].map(String);
  const aIdx = indexMap_(aHdr);

  // Build all rows in memory
  const out = [];
  for (let r = 1; r < aVals.length; r++) {
    const analystId = String(aVals[r][aIdx['analyst_id']] || '').trim();
    if (!analystId) continue;

    const row = computeDailyMetricsRow_(analystId, dateISO);
    out.push([
      row.date,
      row.analyst_id,
      row.available_mins,
      row.handling_mins,
      row.output_total,
      row.standard_mins,
      row.efficiency_pct,
      row.utilisation_pct,
      row.throughput_per_hr,
      row.flags,
      row.notes
    ]);
  }

  if (out.length) {
    const start = dm.getLastRow() + 1;
    dm.getRange(start, 1, out.length, 11).setValues(out);
  }

  return { ok: true, date: dateISO, built_for: out.length };
}

function computeLiveProductionToday_(analystId) {
  const ss = master_();
  const dateISO = toISODate_(new Date());

  // --- Baseline & meetings
  const baselineHours = Number(getBaselineHoursForAnalyst_(analystId) || 8.5);
  const baselineMins = Math.round(baselineHours * 60);
  const meetingMins = getAcceptedMeetingMinutes_(analystId, dateISO);

  // --- Status stints → minutes by state today
  const sl = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
            .filter(r => r.date_str===dateISO && r.analyst_id_norm===normId_(analystId))
            .sort((a,b)=> a.ts - b.ts);
  const stints = stitchStints_(sl);
  const minsByState = {};
  stints.forEach(s => {
    const m = minutesBetween_(s.start, s.end);
    minsByState[s.state] = (minsByState[s.state]||0) + m;
  });

  // Logged-in minutes = everything except LoggedOut
  const loggedInMins = Object.entries(minsByState)
      .filter(([state]) => String(state) !== 'LoggedOut')
      .reduce((a,[,m]) => a + (Number(m)||0), 0);

  // --- Checks today
  const ce = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
            .filter(r => r.date_str===dateISO && r.analyst_id_norm===normId_(analystId));
  const handlingMins = ce.reduce((a,r)=> a + (Number(r.duration_mins)||0), 0);
  const outputTotal = ce.length;

  // Standard mins from CheckTypes × counts
  const types = readCheckTypes_(); // [{name, avg_minutes}]
  const avg = {}; types.forEach(t => avg[t.name] = Number(t.avg_minutes||0));
  const countByType = {};
  ce.forEach(r => { const t=String(r.check_type||''); if (t) countByType[t]=(countByType[t]||0)+1; });
  const standardMins = Object.keys(countByType).reduce((sum,t)=> sum + (countByType[t]* (avg[t]||0)), 0);

  // Available mins = baseline − meetings (≥0)
  const availableMins = Math.max(0, baselineMins - meetingMins);

  // KPIs (guard divisions)
  const live_eff = standardMins > 0 ? Math.round((handlingMins/standardMins)*100) : 0;
  const live_util = availableMins > 0 ? Math.round((handlingMins/availableMins)*100) : 0;
  const live_tph = availableMins > 0 ? Number((outputTotal / (availableMins/60)).toFixed(2)) : 0;

  return {
    dateISO,
    logged_in_mins: loggedInMins,
    live_efficiency_pct: live_eff,
    live_utilisation_pct: live_util,
    live_throughput_per_hr: live_tph
  };
}

// Single source of truth: compute → write KPIs to Live
function updateLiveKPIsFor_(analystId) {
  const k = computeLiveProductionToday_(analystId); // your updated, capped version
  upsertLive_(analystId, {
    logged_in_mins: k.logged_in_mins,
    live_efficiency_pct: k.live_efficiency_pct,
    live_utilisation_pct: k.live_utilisation_pct,
    live_throughput_per_hr: k.live_throughput_per_hr
  });
}
