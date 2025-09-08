/******************************************************
 * 05_checks.gs — Check logging & validation (optimized)
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_, indexMap_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), readRows_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_() (03_sessions.gs)
 * - getCurrentStateInfo_(), refreshLiveFor_() (04_state_engine.gs)
 * - updateLiveKPIsFor_() (07_metrics.gs)
 ******************************************************/

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
  const sh = getOrCreateMasterSheet_(SHEETS.CHECK_EVENTS, HEADERS);

  // If sheet already existed but header order drifted, realign the first row.
  const current = sh.getRange(1, 1, 1, Math.max(HEADERS.length, sh.getLastColumn()))
                    .getValues()[0].map(String);
  HEADERS.forEach((h, i) => { if ((current[i] || '').trim() !== h) sh.getRange(1, i + 1).setValue(h); });
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Given an analyst and date, return the last known state for that day.
 * Falls back to 'Idle' if we have no logs.
 */
function lastStateFor_(analystId, dateISO) {
  const ss = master_();
  const rows = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === dateISO && r.analyst_id_norm === normId_(analystId))
    .sort((a, b) => a.ts - b.ts);

  const last = rows[rows.length - 1];
  return last ? String(last.state || 'Idle') : 'Idle';
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

  // TextFinder can produce false positives (substring matches).
  // Verify exact matches case-insensitively.
  const hit = tf.findNext();
  if (!hit) return; // no textual hit → unique

  // We only need to confirm collision if an exact (trim, lower) match exists.
  // If you want to be extra safe, scan only the handful of candidate cells in the column.
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
 *
 * @param {string} token Session token from the UI
 * @param {string} check_type Must exist in CheckTypes (we do not enforce existence here — UX shows a dropdown)
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
    requireSession_(token);
    const type = String(check_type || '').trim();
    if (!type) throw new Error('Check Type is required.');
    const uid = String(case_id || '').trim();
    if (!uid) throw new Error('Case ID (MP scorecard UID) is required.');
    const dur = Number(duration_mins);
    if (!dur || dur <= 0) throw new Error('Please enter a positive duration (minutes).');

    validateCaseIdUnique_(uid);

    const ts = new Date();
    const dateISO = toISODate_(ts);
    const id = getCurrentAnalystId_();
    const stateAtLog = lastStateFor_(id, dateISO);

    const sh = getOrCreateMasterSheet_(SHEETS.CHECK_EVENTS, [
      'completed_at_iso','date','analyst_id','check_type','case_id',
      'duration_mins','result','rework_flag','state_at_log'
    ]);
    sh.appendRow([ts.toISOString(), dateISO, id, type, uid, dur, String(result||''), String(rework_flag||''), stateAtLog]);

    refreshLiveFor_(id);
    try { updateLiveKPIsFor_(id); } catch(e) {}
    try {postCheckHooks_(analystId, dateISO);} catch (e) {}
    
    return { ok:true, ts: ts.toISOString(), analyst_id:id, check_type:type, state_at_log:stateAtLog };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
    
   
  }
}
