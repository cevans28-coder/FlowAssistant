/******************************************************
 * 05_checks.gs — Check logging & validation (optimized, v2-aware)
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_, indexMap_, readRows_ (01_utils.gs / 02_master_access.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_() (03_sessions.gs)
 * - getCurrentStateInfo_(), refreshLiveFor_() (04_state_engine.gs)
 * - updateLiveKPIsFor_() (07_metrics.gs)
 * - postCheckHooks_() (15_hooks.gs) ← optional, best-effort
 ******************************************************/

// Explicit tab names we read from:
const STATUS_LOGS_TAB_FOR_STATE = 'StatusLogs_v2'; // v2 source of truth for states
const CHECK_EVENTS_TAB = (typeof SHEETS !== 'undefined' && SHEETS.CHECK_EVENTS) ? SHEETS.CHECK_EVENTS : 'CheckEvents_v2';

/**
 * Ensure and return the canonical CheckEvents sheet.
 * We centralize headers here to avoid accidental column drift.
 */
function getCheckEventsSheet_() {
  const HEADERS = [
    'completed_at_iso', // ISO ts when user logged the check
    'date', // yyyy-MM-dd (TZ)
    'analyst_id', // normalized email
    'check_type', // name from CheckTypes
    'case_id', // MP scorecard UID (globally unique)
    'duration_mins', // positive integer
    'result', // optional
    'rework_flag', // optional
    'state_at_log' // analyst state at the time of logging
  ];
  const sh = getOrCreateMasterSheet_(CHECK_EVENTS_TAB, HEADERS);

  // Realign the first row if order drifted.
  const current = sh.getRange(1, 1, 1, Math.max(HEADERS.length, sh.getLastColumn()))
                    .getValues()[0].map(String);
  HEADERS.forEach((h, i) => {
    if ((current[i] || '').trim() !== h) sh.getRange(1, i + 1).setValue(h);
  });
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Given an analyst and date, return the last known state for that day.
 * Falls back to 'Idle' if we have no logs.
 */
function lastStateFor_(analystId, dateISO) {
  const ss = master_();
  const rows = readRows_(ss.getSheetByName(STATUS_LOGS_TAB_FOR_STATE))
    .filter(r => r.date_str === dateISO && r.analyst_id_norm === normId_(analystId))
    .sort((a, b) => a.ts - b.ts);

  const last = rows[rows.length - 1];
  return last ? String(last.state || 'Idle') : 'Idle';
}

/**
 * Optional: validate check_type exists in CheckTypes.
 * We keep this tolerant (UI dropdown typically enforces it).
 */
function validateCheckType_(check_type) {
  const want = String(check_type || '').trim();
  if (!want) throw new Error('Check Type is required.');
  try {
    const list = (typeof readCheckTypes_ === 'function') ? readCheckTypes_() : [];
    if (!list || !list.length) return; // tolerate if not configured yet
    const ok = list.some(t => String(t.name).trim() === want);
    if (!ok) throw new Error('Unknown Check Type: ' + want);
  } catch (e) {
    // If CheckTypes read fails for any reason, don’t block logging:
    // comment out the next line to be strict.
    // throw e;
  }
}

/**
 * Fast-ish global duplicate guard for Case IDs (case-insensitive).
 * Uses a column-scoped TextFinder on the case_id column for speed
 * and then verifies exact, case-insensitive equality.
 *
 * Throws if the caseId is already present anywhere in CheckEvents.
 */
function validateCaseIdUnique_(caseId) {
  const uid = String(caseId || '').trim();
  if (!uid) throw new Error('Case ID (MP scorecard UID) is required.');

  const sh = getCheckEventsSheet_(); // ensures headers & returns sheet
  if (sh.getLastRow() <= 1) return; // no data yet → OK

  // Find the case_id column once
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = indexMap_(hdr);
  const caseCol = idx['case_id'];
  if (caseCol == null) return; // malformed, but don't block logging

  // Scope TextFinder to case_id column only (faster & safer than whole-sheet)
  const colRange = sh.getRange(2, caseCol + 1, sh.getLastRow() - 1, 1);
  const tf = colRange.createTextFinder(uid);
  tf.matchCase(false).matchEntireCell(false);

  // TextFinder can produce substring matches; verify exact trims in column.
  const values = colRange.getValues().flat().filter(Boolean);
  const want = uid.toLowerCase();
  const dup = values.some(v => String(v).trim().toLowerCase() === want);
  if (dup) throw new Error('This Case ID has already been logged: ' + uid);
}

/**
 * Sanitize and validate duration minutes.
 * Returns a positive integer (≥ 1) or throws a helpful error.
 */
function coerceDurationMins_(value) {
  const n = Math.floor(Number(value));
  if (!isFinite(n) || n <= 0) throw new Error('Please enter a positive duration (minutes).');
  return n;
}

/**
 * Log a completed check.
 *
 * Side effects:
 * - Writes a single row to CheckEvents
 * - Refreshes Live to keep UI & TL data fresh
 * - Delegates KPI updates to updateLiveKPIsFor_() so KPI math is centralized
 * - Optionally runs postCheckHooks_ (rebuild per-type & replace DailyMetrics row)
 *
 * @param {string} token Session token from the UI
 * @param {string} check_type Must exist in CheckTypes (dropdown-enforced in UI)
 * @param {string} case_id Required, globally unique across all analysts & days
 * @param {number} duration_mins Positive minutes
 * @param {string} result Optional
 * @param {string} rework_flag Optional (e.g., 'Y'/'N')
 * @returns {{ok: boolean, ts: string, analyst_id: string, check_type: string, state_at_log: string}}
 */
function completeCheck(token, check_type, case_id, duration_mins, result, rework_flag) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error('Please try again (another update is in progress).');

  try {
    token = requireSessionOrAdopt_(token);

    // Inputs
    validateCheckType_(check_type);
    const type = String(check_type || '').trim();

    const uid = String(case_id || '').trim();
    if (!uid) throw new Error('Case ID (MP scorecard UID) is required.');

    const dur = coerceDurationMins_(duration_mins);

    // Global duplicate guard on case_id
    validateCaseIdUnique_(uid);

    // Timestamps & identity
    const ts = new Date();
    const dateISO = toISODate_(ts);
    const id = getCurrentAnalystId_();
    const stateAtLog = lastStateFor_(id, dateISO);

    // Append to CheckEvents
    const sh = getCheckEventsSheet_();
    sh.appendRow([
      ts.toISOString(), // completed_at_iso
      dateISO, // date
      id, // analyst_id
      type, // check_type
      uid, // case_id
      dur, // duration_mins
      String(result || ''), // result
      String(rework_flag || ''), // rework_flag
      stateAtLog // state_at_log
    ]);

    // Keep Live fresh; KPI math centralized
    refreshLiveFor_(id);
    try { _bustTodaySummaryCache_(id); } catch(e) {}
    try { updateLiveKPIsFor_(id); } catch(e) {}

    // Optional hook for downstream pipelines (only if you’ve defined it)
    try {
      if (typeof postCheckHooks_ === 'function') {
        postCheckHooks_(id, dateISO);
      }
    } catch (e) {
      // Non-fatal by design
    }

    return { ok: true, ts: ts.toISOString(), analyst_id: id, check_type: type, state_at_log: stateAtLog };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ------------------ Small helpers / diagnostics (optional) ------------------ */

/** Quick helper to check if a Case ID exists (case-insensitive). */
function caseIdExists_(caseId) {
  try {
    const sh = getCheckEventsSheet_();
    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const idx = indexMap_(hdr);
    const col = idx['case_id'];
    if (col == null || sh.getLastRow() <= 1) return false;
    const vals = sh.getRange(2, col + 1, sh.getLastRow() - 1, 1).getValues().flat();
    const want = String(caseId || '').trim().toLowerCase();
    return vals.some(v => String(v || '').trim().toLowerCase() === want);
  } catch (e) {
    return false;
  }
}

/** Smoke: append a fake check for the current user (use sparingly). */
function CHECKS_Smoke_AppendFake() {
  const tok = makeToken_(); setSessionTokenFor_(getCurrentAnalystId_(), tok);
  return completeCheck(tok, 'Sample Check Type', 'DEMO-' + Utilities.getUuid().slice(0, 8), 7, 'OK', 'N');
}
