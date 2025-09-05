/******************************************************
 * 08_watchdog.gs — Inactivity rules & gentle nudges
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, minutesBetween_, normId_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), readRows_() (02_master_access.gs)
 * - getCurrentAnalystId_(), getSessionTokenFor_(), logLoginEvent_() (03_sessions.gs)
 * - upsertLive_() (04_state_engine.gs)
 ******************************************************/

/**
 * Watchdog — run every 10 minutes (per-user trigger).
 * Uses last_seen_iso (from heartbeat) to detect inactivity.
 * Writes one StatusLogs row only when a transition is needed.
 *
 * Rules:
 * - Working/Admin : >10m no heartbeat → Idle
 * - Idle : >60m no heartbeat → LoggedOut
 * - Break/Lunch : >60m no heartbeat → LoggedOut
 * - Others (Meeting/Training/Coaching/OOO/etc.): >10m no heartbeat → LoggedOut
 */
function watchdog_heartbeat_10min() {
  const up = PropertiesService.getUserProperties();
  const last = up.getProperty('last_seen_iso');
  if (!last) return; // no heartbeat ever → nothing to do

  const now = new Date();
  const lastDt = new Date(last);
  const gapMins = minutesBetween_(lastDt, now);
  if (gapMins <= 10) return; // still fresh

  const id = getCurrentAnalystId_();
  const today = toISODate_(now);
  const ss = master_();

  // Today’s last state record for this analyst
  const statusRows = readRows_(ss.getSheetByName(SHEETS.STATUS_LOGS))
    .filter(r => r.date_str === today && r.analyst_id_norm === id)
    .sort((a, b) => a.ts - b.ts);

  const lastRow = statusRows[statusRows.length - 1] || null;
  const lastState = lastRow ? String(lastRow.state || '') : 'Idle';
  const lastSrc = lastRow ? String(lastRow.source || '').toLowerCase() : '';

  // Avoid duplicate spam if the last was already written by watchdog,
  // except we still allow the Idle→LoggedOut escalation after >60m.
  if (lastRow && lastSrc.indexOf('watchdog') !== -1) {
    if (!(lastState === 'Idle' && gapMins > 60)) return;
  }

  const sh = getOrCreateMasterSheet_(SHEETS.STATUS_LOGS,
    ['timestamp_iso','date','analyst_id','state','source','note']);

  // --- Rules ---

  // Break / Lunch: keep until >60m, then LoggedOut
  if (['Break', 'Lunch'].includes(lastState)) {
    if (gapMins <= 60) return;
    sh.appendRow([now.toISOString(), today, id, 'LoggedOut', 'watchdog', 'Break/Lunch >60m no heartbeat']);
    upsertLive_(id, {
      analyst_id: id,
      online: 'NO',
      last_seen_iso: now.toISOString(),
      state: 'LoggedOut',
      since_iso: now.toISOString()
    });
    logLoginEvent_('WatchdogLogout', 'Break/Lunch >60m no heartbeat', getSessionTokenFor_(id));
    return;
  }

  // Working / Admin: after >10m, soften to Idle
  if (['Working', 'Admin'].includes(lastState)) {
    sh.appendRow([now.toISOString(), today, id, 'Idle', 'watchdog', 'Auto-idle: no heartbeat >10m']);
    upsertLive_(id, {
      analyst_id: id,
      online: 'NO',
      last_seen_iso: now.toISOString(),
      state: 'Idle',
      since_iso: now.toISOString()
    });
    logLoginEvent_('WatchdogIdle', 'Auto-idle: no heartbeat >10m', getSessionTokenFor_(id));
    return;
  }

  // Idle: >60m → LoggedOut
  if (lastState === 'Idle') {
    if (gapMins > 60) {
      sh.appendRow([now.toISOString(), today, id, 'LoggedOut', 'watchdog', 'Idle >60m with no heartbeat']);
      upsertLive_(id, {
        analyst_id: id,
        online: 'NO',
        last_seen_iso: now.toISOString(),
        state: 'LoggedOut',
        since_iso: now.toISOString()
      });
      logLoginEvent_('WatchdogLogout', 'Idle >60m with no heartbeat', getSessionTokenFor_(id));
    }
    return;
  }

  // All other states (Meeting/Training/Coaching/OOO/…): >10m → LoggedOut
  sh.appendRow([now.toISOString(), today, id, 'LoggedOut', 'watchdog', 'No heartbeat >10m']);
  upsertLive_(id, {
    analyst_id: id,
    online: 'NO',
    last_seen_iso: now.toISOString(),
    state: 'LoggedOut',
    since_iso: now.toISOString()
  });
  logLoginEvent_('WatchdogLogout', 'No heartbeat >10m', getSessionTokenFor_(id));
}

/**
 * Optional nudge: if you are currently in a meeting (from CAL_PULL)
 * but your state is still Working, send yourself a reminder email.
 * Call this from your 10-minute cycle if you want gentle prompts.
 */
function nudgeIfInMeetingButWorking_() {
  const now = new Date();
  const dateISO = toISODate_(now);
  const id = getCurrentAnalystId_();
  const ss = master_();

  const calRows = readRows_(ss.getSheetByName(SHEETS.CAL_PULL))
    .filter(r => r.date_str === dateISO && r.analyst_id_norm === id &&
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
        subject: 'Flow Assistant: You’re in a meeting — update your state?',
        htmlBody: 'You seem to be in a meeting now but your state is <b>Working</b>. Please switch to <b>Meeting</b>.'
      });
    } catch (e) { /* ignore mail errors */ }
  }
}
