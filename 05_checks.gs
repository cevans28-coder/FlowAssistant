/******************************************************
 * 05_checks.gs — Check logging & validation
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), readRows_(), indexMap_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_() (03_sessions.gs)
 * - getCurrentStateInfo_(), refreshLiveFor_() (04_state_engine.gs)
 ******************************************************/

/**
 * Return today's last known state for a given analyst (or 'Idle' if none).
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
 * Global duplicate guard for Case IDs (case-insensitive).
 * Throws if the caseId is already present anywhere in CheckEvents.
 */
function validateCaseIdUnique_(caseId) {
  const uid = String(caseId || '').trim();
  if (!uid) throw new Error('Case ID (MP scorecard UID) is required.');

  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.CHECK_EVENTS);
  if (!sh || sh.getLastRow() <= 1) return; // no data yet → ok

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);
  const want = uid.toLowerCase();

  if (idx['case_id'] === undefined) return; // malformed sheet; let the write proceed

  for (let r = 1; r < vals.length; r++) {
    const seen = String(vals[r][idx['case_id']] || '').trim().toLowerCase();
    if (!seen) continue;
    if (seen === want) {
      throw new Error('This Case ID has already been logged: ' + uid);
    }
  }
}

/**
 * Log a completed check.
 * - Requires session
 * - Validates required fields
 * - Global duplicate Case ID guard
 * - Captures current state as state_at_log
 * - Appends to CheckEvents and refreshes Live
 *
 * @param {string} token Session token from UI
 * @param {string} check_type Name from CheckTypes
 * @param {string} case_id MP scorecard UID (required, unique globally)
 * @param {number} duration_mins Positive minutes (required)
 * @param {string} result Optional
 * @param {string} rework_flag Optional (e.g., 'Y'/'N' or '')
 */
function completeCheck(token, check_type, case_id, duration_mins, result, rework_flag) {
  // Session & inputs
  requireSession_(token);

  const type = String(check_type || '').trim();
  if (!type) throw new Error('Check Type is required.');

  const uid = String(case_id || '').trim();
  if (!uid) throw new Error('Case ID (MP scorecard UID) is required.');

  const dur = Number(duration_mins);
  if (!dur || dur <= 0) throw new Error('Please enter a positive duration (minutes).');

  // Duplicate guard (global)
  validateCaseIdUnique_(uid);

  // Compose record
  const ts = new Date();
  const dateISO = toISODate_(ts);
  const id = getCurrentAnalystId_();

  // Capture current state at time of logging
  // (We use StatusLogs for today; if none, default to Idle)
  const stateAtLog = lastStateFor_(id, dateISO);

  // Append to CheckEvents
  const sh = getOrCreateMasterSheet_(SHEETS.CHECK_EVENTS, [
    'completed_at_iso','date','analyst_id','check_type','case_id',
    'duration_mins','result','rework_flag','state_at_log'
    ,'logged_in_mins','live_efficiency_pct','live_utilisation_pct',
  'live_throughput_per_hr']);

  sh.appendRow([
    ts.toISOString(),
    dateISO,
    id,
    type,
    uid,
    dur,
    String(result || ''),
    String(rework_flag || ''),
    stateAtLog
  ]);

  // Keep Live counters fresh
  refreshLiveFor_(id);
try { updateLiveKPIsFor_(id); } catch(e) {}
  return {
    ok: true,
    ts: ts.toISOString(),
    analyst_id: id,
    check_type: type,
    state_at_log: stateAtLog
  };
}
