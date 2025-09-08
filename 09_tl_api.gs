/******************************************************
 * 09_tl_api.gs — Team Lead actions & snapshot
 * Depends on:
 * - TZ, SHEETS, TL_ALLOWED_STATES (00_constants.gs)
 * - normId_, toISODate_, indexMap_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), getMyEmail_() (03_sessions.gs)
 * - upsertLive_() (04_state_engine.gs)
 ******************************************************/

/** Simple auth: is the current user a Team Lead? */
function isTeamLead_() {
  const ss = master_();

  // 1) Settings: TEAM_LEADS = "a@x.com, b@y.com"
  const set = ss.getSheetByName(SHEETS.SETTINGS);
  const me = (getMyEmail_() || '').trim().toLowerCase();
  if (set && set.getLastRow() > 1) {
    const v = set.getDataRange().getValues();
    for (let r = 1; r < v.length; r++) {
      const k = String(v[r][0] || '').trim().toUpperCase();
      if (k === 'TEAM_LEADS') {
        const list = String(v[r][1] || '')
          .toLowerCase()
          .split(/[,\s]+/)
          .map(s => s.trim())
          .filter(Boolean);
        if (list.includes(me)) return true;
      }
    }
  }

  // 2) Analysts.manager: if this user is listed as a manager for anyone
  const a = ss.getSheetByName(SHEETS.ANALYSTS);
  if (a && a.getLastRow() > 1) {
    const vals = a.getDataRange().getValues();
    const idx = indexMap_(vals[0].map(String));
    for (let r = 1; r < vals.length; r++) {
      const mgr = String(vals[r][idx['manager']] || '').toLowerCase().trim();
      if (mgr && mgr === me) return true;
      // Allow "Name <email>" formats
      const m2 = mgr.match(/<([^>]+)>/);
      if (m2 && String(m2[1] || '').toLowerCase().trim() === me) return true;
    }
  }

  return false;
}

/** Guard wrapper for TL-only endpoints. */
function requireTL_() {
  if (!isTeamLead_()) throw new Error('Not authorised');
}

/** Validate a TL-settable state against hard-coded allowlist. */
function stateAllowedForTL_(state) {
  return TL_ALLOWED_STATES.indexOf(String(state || '').trim()) !== -1;
}

/**
 * TL action: set an analyst's state (audited).
 * - Allowed states are hard-coded (TL_ALLOWED_STATES)
 * - Writes StatusLogs with source "manager:<email>"
 * - Updates Live (and clears token when setting LoggedOut)
 * - Invalidates TL snapshot cache for "today" so changes are immediate in the console
 */
function tlSetState(analystId, newState, note, managerEmail) {
  requireTL_();

  const id = String(analystId || '').trim();
  if (!id) throw new Error('Missing analystId');

  const state = String(newState || '').trim();
  if (!stateAllowedForTL_(state)) {
    throw new Error('State not allowed for TL: ' + state);
  }

  const now = new Date();
  const today = toISODate_(now);

  // 1) Audit in StatusLogs
  const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS,
    ['timestamp_iso','date','analyst_id','state','source','note']);
  const mgr = (managerEmail || getMyEmail_() || 'unknown').toString();
  sh.appendRow([now.toISOString(), today, id, state, `manager:${mgr}`, note || '']);

  // 2) Update Live (and revoke token for LoggedOut)
  const patch = {
    state,
    since_iso: now.toISOString(),
    last_seen_iso: now.toISOString()
  };
  if (state === 'LoggedOut') patch.session_token = ''; // revoke sessions
  upsertLive_(id, patch);

  // 3) Bust TL snapshot cache for today
  invalidateTLSnapshotCacheFor_(today);

  return { ok: true, analyst_id: id, state, ts: now.toISOString() };
}

/** Convenience: TL force-logout. */
function tlForceLogout(analystId, note, managerEmail) {
  return tlSetState(analystId, 'LoggedOut', note, managerEmail);
}

/* -------------------------------------------------------
 * TL Console snapshot — public entry (cached ~45 seconds)
 * ------------------------------------------------------- */

/**
 * Public: Returns TL snapshot for a given date (any of several formats).
 * - Tries script cache first (per-day key)
 * - Falls back to compute helper and caches the result
 */
function getTLSnapshotData(dateISO) {
  const day = normaliseToISODate_(dateISO, TZ);
  const key = 'TLSNAP:' + day;
  const cache = CacheService.getScriptCache();

  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* ignore corrupt cache */ }
  }

  const out = _computeTLSnapshotData_(day);
  try { cache.put(key, JSON.stringify(out), 45); } catch(e) {}
  return out;
}

/** Remove the cached snapshot for a given date (YYYY-MM-DD). */
function invalidateTLSnapshotCacheFor_(dateISO) {
  try {
    const key = 'TLSNAP:' + normaliseToISODate_(dateISO, TZ);
    CacheService.getScriptCache().remove(key);
  } catch (e) {}
}

/**
 * Private: heavy-lift snapshot builder.
 * Accepts a normalised 'YYYY-MM-DD' and returns:
 * { rows:[{...}], kpis:{...}, date:'YYYY-MM-DD' }
 *
 * - LIVE is the primary source for presence and *live* KPIs.
 * - DAILY metrics are used as a fallback when live values are blank/zero.
 * - Tolerant header lookups (no hard crashes if columns are missing).
 */
function _computeTLSnapshotData_(dateISO) {
  const ss = master_();
  const dateStr = normaliseToISODate_(dateISO, TZ); // safe

  // ---- LIVE (current snapshot) ----
  const live = ss.getSheetByName(SHEETS.LIVE);
  const rowsOut = [];
  if (live && live.getLastRow() > 1) {
    const lVals = live.getDataRange().getValues();
    const lHdr = lVals[0].map(String);
    const L = indexMap_(lHdr);
    const col = (name) => (Object.prototype.hasOwnProperty.call(L, name) ? L[name] : -1);

    for (let r = 1; r < lVals.length; r++) {
      const v = lVals[r];

      const iAnalyst = col('analyst_id');
      const analyst_id = iAnalyst >= 0 ? String(v[iAnalyst] || '').trim().toLowerCase() : '';
      if (!analyst_id) continue;

      const getS = (n) => { const i = col(n); return i >= 0 ? String(v[i] || '') : ''; };
      const getN = (n) => { const i = col(n); return i >= 0 ? Number(v[i] || 0) || 0 : 0; };
      const getB = (n) => { const i = col(n); return i >= 0 ? (String(v[i] || '').toUpperCase() === 'YES') : false; };

      rowsOut.push({
        analyst_id,
        name: getS('name'),
        team: getS('team'),
        state: getS('state') || 'Idle',
        since_iso: getS('since_iso'),
        online: getB('online'),
        today_checks: getN('today_checks'),
        efficiency_pct: getN('live_efficiency_pct'),
        utilisation_pct: getN('live_utilisation_pct'),
        throughput_per_hr: getN('live_throughput_per_hr'),
        logged_in_mins: getN('logged_in_mins'),
        location: (col('location_today') >= 0 ? getS('location_today') : '')
      });
    }
  }

  // ---- DAILY METRICS (date-filtered; robust match) ----
  const dm = ss.getSheetByName(SHEETS.DAILY);
  const metricsByAnalyst = {};
  if (dm && dm.getLastRow() > 1) {
    const mVals = dm.getDataRange().getValues();
    const mHdr = mVals[0].map(String);
    const M = indexMap_(mHdr);
    const mCol = (name) => (Object.prototype.hasOwnProperty.call(M, name) ? M[name] : -1);

    for (let r = 1; r < mVals.length; r++) {
      const row = mVals[r];

      // Tolerant date match
      const rawDate = mCol('date') >= 0 ? row[mCol('date')] : null;
      const rowISO = normaliseToISODate_(rawDate, TZ);
      if (rowISO !== dateStr) continue;

      const iAnalyst = mCol('analyst_id');
      const aid = iAnalyst >= 0 ? String(row[iAnalyst] || '').trim().toLowerCase() : '';
      if (!aid) continue;

      const toNum = (n) => (mCol(n) >= 0 ? Number(row[mCol(n)] || 0) || 0 : 0);

      metricsByAnalyst[aid] = {
        efficiency_pct: toNum('efficiency_pct'),
        utilisation_pct: toNum('utilisation_pct'),
        throughput_per_hr: toNum('throughput_per_hr')
      };
    }
  }

  // ---- Merge (prefer live KPIs; fallback to daily) ----
  let totalChecks = 0, effSum = 0, utilSum = 0, effCount = 0, utilCount = 0;
  let nWorking = 0, nMeeting = 0, nOther = 0, nLoggedOut = 0;

  rowsOut.forEach(r => {
    totalChecks += r.today_checks || 0;

    // state counters
    const s = (r.state || '').toLowerCase();
    if (s === 'loggedout') nLoggedOut++;
    else if (s === 'working' || s === 'admin') nWorking++;
    else if (s === 'meeting') nMeeting++;
    else nOther++;

    // Prefer live; fallback to daily if live is 0/blank
    const daily = metricsByAnalyst[r.analyst_id] || {};
    if (!r.efficiency_pct && daily.efficiency_pct) r.efficiency_pct = daily.efficiency_pct;
    if (!r.utilisation_pct && daily.utilisation_pct) r.utilisation_pct = daily.utilisation_pct;
    if (!r.throughput_per_hr && daily.throughput_per_hr) r.throughput_per_hr = daily.throughput_per_hr;

    if (typeof r.efficiency_pct === 'number') { effSum += r.efficiency_pct; effCount++; }
    if (typeof r.utilisation_pct === 'number') { utilSum += r.utilisation_pct; utilCount++; }
  });

  const kpis = {
    total_checks: totalChecks,
    avg_efficiency: effCount ? Math.round(effSum / effCount) : 0,
    avg_utilisation: utilCount ? Math.round(utilSum / utilCount) : 0,
    working: nWorking,
    meeting: nMeeting,
    other: nOther,
    loggedout: nLoggedOut
  };

  return { rows: rowsOut, kpis, date: dateStr };
}

/* -------------------------------------------------------
 * Helpers used by TL UI (timeline & date parsing)
 * ------------------------------------------------------- */

/**
 * Normalise many date inputs to 'YYYY-MM-DD' in the given timezone.
 * Handles:
 * - 'YYYY-MM-DD' (returns as-is)
 * - 'DD/MM/YYYY' → converts
 * - 'MM/DD/YYYY' → best-effort (if day > 12 treat as DD/MM)
 * - Date objects → formats
 * - ISO timestamps '2025-09-04T10:20:00Z' → formats
 */
function normaliseToISODate_(input, tz) {
  if (!input) return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Date object
  if (input instanceof Date && !isNaN(input)) {
    return Utilities.formatDate(input, tz, 'yyyy-MM-dd');
  }

  const s = String(input).trim();

  // Already ISO YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // ISO timestamp → Date → ISO
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }

  // DD/MM/YYYY or MM/DD/YYYY
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    let d = parseInt(m[1], 10);
    let mo = parseInt(m[2], 10);
    const y = parseInt(m[3], 10);

    // Disambiguate
    if (d > 12) {
      // DD/MM — fine
    } else if (mo > 12) {
      // impossible month, swap (input was MM/DD but invalid) → treat as DD/MM
      const tmp = d; d = mo; mo = tmp;
    } else {
      // ambiguous like 04/05/2025 → assume UK style (DD/MM)
    }

    const jsDate = new Date(y, mo - 1, d, 12, 0, 0); // Noon to avoid DST edges
    return Utilities.formatDate(jsDate, tz, 'yyyy-MM-dd');
  }

  // Fallback: try Date()
  const tryDate = new Date(s);
  if (!isNaN(tryDate)) {
    return Utilities.formatDate(tryDate, tz, 'yyyy-MM-dd');
  }

  // Last resort: today
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

/**
 * TL: Day timeline for one analyst (read-only).
 * Caps to the selected day and (if today) to the current time.
 */
function getAnalystDayTimeline(analystId, dateISO) {
  if (!analystId) throw new Error('Missing analystId');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateISO)) throw new Error('Use YYYY-MM-DD');

  const ss = master_();
  const idNorm = normId_(analystId);

  // 1) Status → stints, capped to the chosen day
  const sl = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === dateISO && r.analyst_id_norm === idNorm)
    .sort((a,b)=> a.ts - b.ts);

  const { start, end } = computeDayBounds_(dateISO);
  const stints = [];
  for (let i=0;i<sl.length;i++){
    const cur = sl[i], next = sl[i+1];
    if (!cur.ts) continue;
    let s = new Date(Math.max(cur.ts.getTime(), start.getTime()));
    let e = next && next.ts ? new Date(Math.min(next.ts.getTime(), end.getTime())) : new Date(end);
    if (e > s) stints.push({ state: String(cur.state||'Idle'), start_iso: s.toISOString(), end_iso: e.toISOString() });
  }

  // 2) Checks for the date
  const ce = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
    .filter(r => r.date_str === dateISO && r.analyst_id_norm === idNorm)
    .map(r => ({
      ts_iso: r.completed_at_iso || (r.ts ? r.ts.toISOString() : null),
      check_type: String(r.check_type||''),
      case_id: String(r.case_id||''),
      duration_mins: Number(r.duration_mins||0),
      state_at_log: String(r.state_at_log||'')
    }))
    .filter(x => x.ts_iso);

  // 3) Calendar (accepted) for context (optional)
  const cal = readRows_(ss.getSheetByName(SHEETS.CAL_PULL))
    .filter(r => r.date_str === dateISO && r.analyst_id_norm === idNorm && String(r.my_status).toUpperCase() === 'YES')
    .map(r => ({
      start_iso: r.start_iso, end_iso: r.end_iso, title: String(r.title||''), category: String(r.category||'')
    }));

  return { ok:true, date: dateISO, analyst_id: analystId, stints, checks: ce, calendar: cal };
}

/** Local day bounds helper to avoid cross-file dependency issues. */
function computeDayBounds_(dateISO) {
  const start = Utilities.parseDate(dateISO + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss');
  let end = Utilities.parseDate(dateISO + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');

  // If the requested date is "today" (in TZ), cap end at now.
  const todayISO = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (dateISO === todayISO) {
    const now = new Date();
    if (now < end) end = now;
  }
  // Guard: if end somehow ends up <= start, nudge it to start+1min
  if (end <= start) end = new Date(start.getTime() + 60 * 1000);

  return { start, end };
}
