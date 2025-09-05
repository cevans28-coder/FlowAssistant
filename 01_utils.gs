/******************************************************
 * Flow Assistant — Utilities
 * Shared helper functions used across the project
 ******************************************************/

/** Return a mapping of { headerName: columnIndex } */
function indexMap_(headers) {
  const m = {};
  (headers || []).forEach((h, i) => m[String(h).trim()] = i);
  return m;
}

/** Normalise analyst IDs and emails (lowercase + trim) */
function normId_(s) {
  return String(s || '').toLowerCase().trim();
}

/** Format a Date to ISO string (yyyy-MM-dd) in project TZ */
function toISODate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/** Minutes between two Date objects */
function minutesBetween_(a, b) {
  return Math.max(0, Math.round((b - a) / 60000));
}

/** Rounding helpers */
function round0(n) { return Math.round(Number(n) || 0); }
function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function roundPct(n) { return Math.round((Number(n) || 0) * 10000) / 10000; }

/**
 * Convert a sheet into array of row objects with:
 * - normalised fields
 * - parsed timestamps
 * - date_str field
 */
function readRows_(sh) {
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length <= 1) return [];
  const hdr = vals[0].map(h => String(h).trim());
  const out = [];

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (!row || !row.some(c => c !== '')) continue;

    const o = {};
    hdr.forEach((h, i) => o[h] = row[i]);

    // Normalise analyst_id
    o.analyst_id_norm = normId_(o['analyst_id']);

    // Parse timestamps
    const tsIso = o['timestamp_iso'] || o['completed_at_iso'] || null;
    if (tsIso) {
      const ts = new Date(tsIso);
      if (!isNaN(ts)) o.ts = ts;
    }
    if (o['start_iso']) o.start = new Date(o['start_iso']);
    if (o['end_iso']) o.end = new Date(o['end_iso']);

    // Normalised date_str
    let dateStr = '';
    if (o.ts instanceof Date && !isNaN(o.ts)) {
      dateStr = Utilities.formatDate(o.ts, TZ, 'yyyy-MM-dd');
    } else {
      const raw = o['date'];
      if (typeof raw === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(raw.trim())) {
        dateStr = raw.trim();
      } else if (raw instanceof Date && !isNaN(raw)) {
        dateStr = Utilities.formatDate(raw, TZ, 'yyyy-MM-dd');
      } else {
        dateStr = String(raw || '').trim();
      }
    }
    o.date_str = dateStr;

    out.push(o);
  }
  return out;
}

/**
 * Delete rows matching predicate
 * @param {Sheet} sh - target sheet
 * @param {function(Array):boolean} pred - returns true to delete row
 */
function deleteRowsBy_(sh, pred) {
  if (!sh) return;
  const v = sh.getDataRange().getValues();
  if (!v || v.length <= 1) return;
  const hdr = v[0];
  const keep = v.slice(1).filter(r => !pred(r));

  sh.clearContents();
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  if (keep.length) sh.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
}

// Sum minutes today where state != 'LoggedOut', capping ongoing stint at NOW
function computeLoggedInMinutesToday_(analystId) {
  const ss = master_();
  const now = new Date();
  const today = toISODate_(now); // e.g., '2025-09-03'

  // Normalised rows for TODAY only, sorted
  const rows = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === normId_(analystId))
    .sort((a,b) => a.ts - b.ts);

  if (!rows.length) return 0;

  // Build stints: each log starts a stint; end is next log, or end-of-day — but capped at NOW
  const endOfDay = Utilities.parseDate(today + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');
  let total = 0;

  for (let i = 0; i < rows.length; i++) {
    const cur = rows[i];
    const next = rows[i+1];

    const start = cur.ts;
    if (!start) continue;

    // Natural end (next log or end-of-day)
    let end = next && next.ts ? next.ts : endOfDay;

    // Cap at NOW for live counting
    if (end > now) end = now;

    // Skip invalid/zero stints
    if (!end || end <= start) continue;

    // Exclude LoggedOut stints
    const state = String(cur.state || '');
    if (state === 'LoggedOut') continue;

    total += minutesBetween_(start, end); // your existing helper
  }

  // Guard against negative/NaN
  return Math.max(0, Math.round(total));
}
// Compute TODAY handling_mins, output_total, and standard_mins so far
function computeLiveProductionToday_(analystId) {
  const ss = master_();
  const dateISO = toISODate_(new Date());

  // Baseline & meetings
  const baselineHours = Number(getBaselineHoursForAnalyst_(analystId) || 8.5);
  const baselineMins = Math.max(0, Math.round(baselineHours * 60));
  const meetingMins = Math.max(0, getAcceptedMeetingMinutes_(analystId, dateISO));
  const availableMins = Math.max(0, baselineMins - meetingMins);

  // Status stints today (cap open stint at NOW)
  const sl = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
            .filter(r => r.date_str === dateISO && r.analyst_id_norm === normId_(analystId))
            .sort((a,b)=> a.ts - b.ts);
  const stints = stitchStintsCappedToNow_(sl, dateISO);

  // Logged-in = Working + Idle
  const LOGGED_IN_STATES = new Set(['Working','Idle']);
  let loggedInMins = 0;
  for (const s of stints) {
    if (LOGGED_IN_STATES.has(String(s.state))) {
      loggedInMins += minutesBetween_(s.start, s.end);
    }
  }
  loggedInMins = Math.max(0, Math.round(loggedInMins));

  // Checks today
  const ce = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
            .filter(r => r.date_str === dateISO && r.analyst_id_norm === normId_(analystId));
  const handlingMins = Math.max(0, Math.round(ce.reduce((a,r)=> a + (Number(r.duration_mins)||0), 0)));
  const outputTotal = ce.length;

  // Standard mins (CheckTypes × counts)
  const types = readCheckTypes_(); const avg = {};
  types.forEach(t => avg[t.name] = Number(t.avg_minutes || 0));
  const countByType = {};
  ce.forEach(r => { const t = String(r.check_type || ''); if (t) countByType[t] = (countByType[t]||0) + 1; });
  const standardMins = Object.keys(countByType).reduce((sum,t)=> sum + (countByType[t] * (avg[t]||0)), 0);

  // KPIs (no caps; guard denominators)
  const live_eff = standardMins > 0 ? Math.round((handlingMins / standardMins) * 100) : 0;
  const live_util = availableMins > 0 ? Math.round((handlingMins / availableMins) * 100) : 0;
  const live_tph = availableMins > 0 ? Number((outputTotal / (availableMins/60)).toFixed(2)) : 0;

  return {
    dateISO,
    logged_in_mins: loggedInMins,
    live_efficiency_pct: live_eff,
    live_utilisation_pct: live_util, // can be >100 by design now
    live_throughput_per_hr: live_tph
  };
}

function updateLiveKPIsFor_(analystId) {
  const id = normId_(analystId || getCurrentAnalystId_());
  const ss = master_();

  // Pull inputs (today only)
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  // Baseline hours
  const baselineHrs = Number(getBaselineHoursForAnalyst_(id) || 8.5);
  const baselineMins = Math.max(0, Math.round(baselineHrs * 60));

  // Meetings (accepted)
  const meetMins = getAcceptedMeetingMinutes_(id, today) || 0;

  // Available mins (cannot be negative)
  const availableMins = Math.max(0, baselineMins - meetMins);

  // Handling + output + standard mins from CheckEvents
  let handlingMins = 0, outputTotal = 0, standardMins = 0;
  const avg = {};
  (readCheckTypes_() || []).forEach(t => avg[String(t.name)] = Number(t.avg_minutes || 0));

  const ce = ss.getSheetByName(SHEETS.CHECK_EVENTS);
  if (ce && ce.getLastRow() > 1) {
    const vals = ce.getDataRange().getValues();
    const idx = indexMap_(vals[0].map(String));
    for (let r=1;r<vals.length;r++){
      const row = vals[r];
      if (String(row[idx['date']]||'') !== today) continue;
      if (normId_(row[idx['analyst_id']]) !== id) continue;
      const mins = Number(row[idx['duration_mins']] || 0);
      const ct = String(row[idx['check_type']] || '');
      if (mins > 0) handlingMins += mins;
      if (ct) {
        outputTotal += 1;
        standardMins += (avg[ct] || 0);
      }
    }
  }

  // Logged-in mins (Working/Admin/Meeting/Training/Coaching/Break/Lunch/Idle — i.e., everything except LoggedOut/OOO)
  const sl = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a,b)=> a.ts - b.ts);

  const { start, end } = (typeof computeDayBounds_ === 'function') 
    ? computeDayBounds_(today) 
    : { 
        start: Utilities.parseDate(today+' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss'), 
        end: new Date() 
      };

  let loggedInMins = 0;
  for (let i=0;i<sl.length;i++){
    const cur = sl[i], next = sl[i+1];
    if (!cur.ts) continue;
    const s = new Date(Math.max(cur.ts.getTime(), start.getTime()));
    const e = next && next.ts ? new Date(Math.min(next.ts.getTime(), end.getTime())) : new Date(end);
    if (e <= s) continue;
    const st = String(cur.state||'');
    if (st && !/^(LoggedOut|OOO)$/i.test(st)) {
      loggedInMins += Math.round((e - s)/60000);
    }
  }

  // KPIs with guards (cap utilisation at 200% to avoid wild test loops; you can choose 100 if you prefer)
  const eff = (standardMins > 0) ? Math.round((handlingMins / standardMins) * 100) : 0;
  const utlRaw = (availableMins > 0) ? (handlingMins / availableMins) * 100 : 0;
  const utl = Math.min(200, Math.round(utlRaw)); // or 100 to hard-cap
  const tput = (availableMins > 0) ? Number((outputTotal / (availableMins / 60)).toFixed(2)) : 0;

  // Write to Live (numeric only, no blanks)
  const live = ss.getSheetByName(SHEETS.LIVE);
  if (!live || live.getLastRow() < 2) return;
  const lVals = live.getDataRange().getValues();
  const L = indexMap_(lVals[0].map(String));
  let rowIndex = -1;
  for (let r=1;r<lVals.length;r++){
    if (normId_(lVals[r][L['analyst_id']]) === id) { rowIndex = r+1; break; }
  }
  if (rowIndex === -1) return;

  const safeSet = (colName, value) => {
    if (L[colName] != null) live.getRange(rowIndex, L[colName]+1).setValue(Number(value)||0);
  };
  safeSet('logged_in_mins', loggedInMins);
  safeSet('live_efficiency_pct', eff);
  safeSet('live_utilisation_pct', utl);
  safeSet('live_throughput_per_hr', tput);
}
