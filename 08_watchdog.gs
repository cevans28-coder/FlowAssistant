/******************************************************
 * 08_watchdog.gs — Inactivity rules & gentle nudges (optimized)
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, minutesBetween_, normId_, readRows_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), getSessionTokenFor_(), logLoginEvent_() (03_sessions.gs)
 * - upsertLive_() (04_state_engine.gs)
 *
 * Triggers (installed by ensureUserTriggers_ in 99_triggers.gs):
 * - watchdog_heartbeat_10min() — every 10 minutes per user
 *
 * Behaviour:
 * - If no heartbeat in >10m while Working/Admin → set Idle
 * - If Idle and >60m without heartbeat → LoggedOut
 * - If Break/Lunch and >60m without heartbeat → LoggedOut
 * - All other states (Meeting/Training/Coaching/OOO/…) with >10m gap → LoggedOut
 *
 * Notes:
 * - We write ONE StatusLogs row only when a transition is required.
 * - We keep Live in sync and add an audit in LoginHistory.
 ******************************************************/

/**
 * Watchdog — per-user trigger every ~10m.
 * Uses last_seen_iso from PropertiesService(User) (set by heartbeat()).
 * Writes to StatusLogs only if a state transition is needed.
 */
function watchdog_heartbeat_10min() {
  // 0) Read last heartbeat. If never set, we can’t infer inactivity → skip.
  const up = PropertiesService.getUserProperties();
  const last = up.getProperty('last_seen_iso');
  if (!last) return;

  const now = new Date();
  const lastDt = new Date(last);
  const gap = minutesBetween_(lastDt, now);
  if (gap <= 10) return; // still fresh; nothing to do

  const id = getCurrentAnalystId_();
  const today = toISODate_(now);
  const ss = master_();

  // 1) Read today’s last known state for this user
  const statusRows = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);

  const lastRow = statusRows[statusRows.length - 1] || null;
  const lastState = lastRow ? String(lastRow.state || '') : 'Idle';
  const lastSrc = (lastRow ? String(lastRow.source || '') : '').toLowerCase();

  // If the last write came from watchdog already, avoid spamming,
  // but still allow the Idle→LoggedOut escalation (>60m)
  if (lastRow && lastSrc.indexOf('watchdog') !== -1) {
    if (!(lastState === 'Idle' && gap > 60)) return;
  }

  // Helper to append to StatusLogs + LIVE + audit in one shot
  function writeTransition_(newState, note) {
    const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS,
      ['timestamp_iso','date','analyst_id','state','source','note']);

    sh.appendRow([now.toISOString(), today, id, newState, 'watchdog', note]);

    // Keep Live view consistent
    upsertLive_(id, {
      analyst_id: id,
      online: (newState === 'LoggedOut') ? 'NO' : 'NO', // when watchdog mutates, mark NO
      last_seen_iso: now.toISOString(),
      state: newState,
      since_iso: now.toISOString()
    });

    // Audit trail for session events
    const token = getSessionTokenFor_(id);
    logLoginEvent_(
      newState === 'Idle' ? 'WatchdogIdle' : 'WatchdogLogout',
      note,
      token
    );
  }

  /* =================== RULES =================== */

  // Break/Lunch: allow up to 60m without heartbeat, then logout
  if (lastState === 'Break' || lastState === 'Lunch') {
    if (gap > 60) writeTransition_('LoggedOut', 'Break/Lunch >60m with no heartbeat');
    return;
  }

  // Working/Admin: after >10m without heartbeat → auto-Idle (softer)
  if (lastState === 'Working' || lastState === 'Admin') {
    writeTransition_('Idle', 'Auto-idle: no heartbeat >10m');
    return;
  }

  // Idle: after >60m → LoggedOut
  if (lastState === 'Idle') {
    if (gap > 60) writeTransition_('LoggedOut', 'Idle >60m with no heartbeat');
    return;
  }

  // Other states (Meeting/Training/Coaching/OOO/…):
  // If no heartbeat for >10m, assume they are gone → LoggedOut
  writeTransition_('LoggedOut', 'No heartbeat >10m');
}

/**
 * Optional nudge:
 * If you are currently in a (accepted) calendar meeting but your state is Working,
 * send yourself a polite email reminder to update your state.
 * Safe to call from the 10-minute cycle; errors are swallowed.
 */
function nudgeIfInMeetingButWorking_() {
  try {
    const now = new Date();
    const dateISO = toISODate_(now);
    const id = getCurrentAnalystId_();
    const ss = master_();

    const calRows = readRows_(ss.getSheetByName(SHEETS.CAL_PULL))
      .filter(r => r.date_str === dateISO &&
                   r.analyst_id_norm === id &&
                   String(r.my_status || '').toLowerCase() === 'accepted');

    const inMeeting = calRows.some(r => r.start && r.end && now >= r.start && now <= r.end);
    if (!inMeeting) return;

    const statusRows = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
      .filter(r => r.date_str === dateISO && r.analyst_id_norm === id)
      .sort((a, b) => a.ts - b.ts);

    const last = statusRows[statusRows.length - 1];
    if (last && String(last.state) === 'Working') {
      const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
      if (!email) return;
      try {
        MailApp.sendEmail({
          to: email,
          subject: 'Flow Assistant — quick reminder',
          htmlBody: 'We detected you appear to be in a meeting, but your state is <b>Working</b>. Want to switch to <b>Meeting</b>?'
        });
      } catch (e) { /* ignore email errors */ }
    }
  } catch (e) { /* swallow to keep trigger healthy */ }
}
