/******************************************************
 * 99_triggers.gs — Menus, time-based triggers & A1 filters
 * Depends on:
 * - TZ, MASTER_SPREADSHEET_ID, SHEETS, DATE_FILTER_SHEETS (00_constants.gs)
 * - toISODate_, indexMap_ (01_utils.gs)
 * - master_() (02_master_access.gs)
 * - getCurrentAnalystId_() (03_sessions.gs)
 * - refreshLiveFor_() (04_state_engine.gs)
 * - syncMyCalendarForDate(), syncMyCalendarTodayNoToken() (06_calendar.gs)
 ******************************************************/

/* =================== Spreadsheet menu =================== */
function onOpen() {
  try { SpreadsheetApp.getActive().setSpreadsheetTimeZone(TZ); } catch (e) {}

  SpreadsheetApp.getUi()
    .createMenu('Flow Assistant')
    .addItem('Open Flow Assistant', 'openDock')
    .addSeparator()
    .addItem('Setup / Repair Master Tabs', 'remoteSetupMaster') // reuse your existing setup
    .addSeparator()
    .addItem('Sync MY Calendar (Today → MASTER)', 'syncMyCalendarTodayNoToken')
    .addItem('Build My Metrics (Today → MASTER)', 'buildMyMetricsToday')
    .addSeparator()
    .addItem('Install/Refresh Triggers (me)', 'ensureUserTriggers_')
    .addSeparator()
    .addItem('Force Logout (admin)…', 'openForceLogoutDialog') // NEW
    .addToUi();
}

// --- Templating helper for <?!= include('...') ?> in HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- Open the panel inside Sheets (modeless)
function openDock() {
  const t = HtmlService.createTemplateFromFile('FlowUI');
  const out = t.evaluate().setTitle('Flow Assistant V1.0').setWidth(520).setHeight(720);
  SpreadsheetApp.getUi().showModelessDialog(out, 'Flow Assistant');
}

// --- Web App entry (if you deploy as web app)
function doGet() {
  var t = HtmlService.createTemplateFromFile('FlowUI');
  return t.evaluate()
           .setTitle('Flow Assistant V1.0')
           .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* =================== Time-based triggers =================== */
/**
 * Removes any duplicates and (re)installs:
 * - autoSyncMyCalendar_10min (every 10 minutes)
 * - watchdog_heartbeat_10min (every 10 minutes)
 * - onMasterEdit_ (installable onEdit on the MASTER workbook for A1 date filters)
 */
function ensureUserTriggers_() {
  const toKeep = new Set(['autoSyncMyCalendar_10min', 'watchdog_heartbeat_10min', 'onMasterEdit_']);
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (toKeep.has(fn)) ScriptApp.deleteTrigger(t);
  });

  // Every 10 minutes: calendar pull + meeting nudge + live refresh
  ScriptApp.newTrigger('autoSyncMyCalendar_10min').timeBased().everyMinutes(10).create();

  // Every 10 minutes: inactivity watchdog
  ScriptApp.newTrigger('watchdog_heartbeat_10min').timeBased().everyMinutes(10).create();

  // Installable onEdit bound to the MASTER spreadsheet (A1 date dropdown filters)
  if (!MASTER_SPREADSHEET_ID || MASTER_SPREADSHEET_ID === 'PASTE_MASTER_ID_HERE') {
    throw new Error('MASTER_SPREADSHEET_ID is not set in 00_constants.gs');
  }
  ScriptApp.newTrigger('onMasterEdit_')
    .forSpreadsheet(MASTER_SPREADSHEET_ID)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('Flow Assistant\n\nTriggers installed/refreshed for your account.');
}

/**
 * Runs every 10 minutes (per user) if you installed ensureUserTriggers_().
 * - Syncs accepted calendar events for today to MASTER
 * - Optionally nudges if you’re in a meeting but state is Working
 * - Refreshes your Live snapshot
 */
function autoSyncMyCalendar_10min() {
  try {
    const today = toISODate_(new Date());
    syncMyCalendarForDate(null, today);
  } catch (e) {
    // swallow errors to keep trigger healthy
  }

  // Optional gentle nudge (only if you kept nudge function in 08_watchdog.gs)
  try { nudgeIfInMeetingButWorking_(); } catch (e) {}

  // Keep Live presence fresh
  try { refreshLiveFor_(getCurrentAnalystId_()); } catch (e) {}
}

/* =================== A1 date dropdown + filter (MASTER) =================== */
/**
 * Installable onEdit handler attached to MASTER.
 * If the edited sheet is one of DATE_FILTER_SHEETS and the cell is A1,
 * apply a text-equals filter on the 'date' column for the selected value.
 */
function onMasterEdit_(e) {
  try {
    if (!e || !e.source) return;
    const sh = e.range && e.range.getSheet ? e.range.getSheet() : e.source.getActiveSheet();
    if (!sh) return;

    if (!Array.isArray(DATE_FILTER_SHEETS) || !DATE_FILTER_SHEETS.includes(sh.getName())) return;
    if (e.range.getA1Notation() !== 'A1') return;

    const hdrVals = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    const idx = indexMap_(hdrVals.map(String));
    if (idx['date'] === undefined) return;

    applyDateFilterForSheet_(sh, idx['date'] + 1, String(e.range.getDisplayValue() || '').trim());
  } catch (err) {
    // no-op; avoid breaking the trigger
  }
}

/** Helper to apply a filter on the 'date' column to a specific value. */
function applyDateFilterForSheet_(sh, dateCol, dateStr) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  if (!sh.getFilter()) {
    try { sh.getRange(1, 1, lastRow, lastCol).createFilter(); } catch (e) {}
  }
  const f = sh.getFilter();
  if (!f) return;

  const crit = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(dateStr || '').build();
  try { f.setColumnFilterCriteria(dateCol, crit); } catch (e) {}
}
/** Show the Force Logout dialog (dropdown of online users) */
function openForceLogoutDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AdminForceLogout')
    .setWidth(380).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Force Logout');
}

/** Return list of analysts considered “logged in” right now. */
function getOnlineAnalysts_() {
  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.LIVE);
  if (!sh || sh.getLastRow() < 2) return [];

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);

  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const state = String(row[idx['state']] || '');
    const token = String(row[idx['session_token']] || '');
    const online = String(row[idx['online']] || '').toUpperCase() === 'YES';

    // “Logged in” if they’re not already LoggedOut and still have a session token
    if (token && state !== 'LoggedOut' && online) {
      out.push({
        analyst_id: String(row[idx['analyst_id']] || ''),
        name: String((idx['name']!=null?row[idx['name']]:'' ) || ''),
        team: String((idx['team']!=null?row[idx['team']]:'' ) || ''),
        state,
        since_iso: String((idx['since_iso']!=null?row[idx['since_iso']]:'' ) || '')
      });
    }
  }
  // Sort by name then analyst_id to keep the dropdown tidy
  out.sort((a,b)=> (a.name||'').localeCompare(b.name||'') || (a.analyst_id||'').localeCompare(b.analyst_id||''));
  return out;
}

/** Force-logout a single analyst (clears token, logs audit, sets state) */
function adminForceLogout(analystId, reasonOpt) {
  if (!analystId) throw new Error('Pick a user to log out.');
  const manager = Session.getActiveUser().getEmail() || 'admin@';
  const note = (reasonOpt || 'Admin force logout') + ` (by ${manager})`;

  // Audit to StatusLogs and Live (reuses your TL state setter)
  tlSetState(analystId, 'LoggedOut', note, manager);

  // Extra audit in LoginHistory (for a clean trail)
  const now = new Date();
  const hist = getOrCreateMasterSheet_(SHEETS.LOGIN_HISTORY,
    ['timestamp_iso','date','analyst_id','event','note','session_token']);
  hist.appendRow([now.toISOString(), toISODate_(now), analystId, 'AdminForceLogout', note, '']);

  return { ok:true, analyst_id:analystId, note };
}

/** Optional: force-logout EVERYONE with an active session */
function adminForceLogoutAll(reasonOpt) {
  const list = getOnlineAnalysts_();
  const msg = reasonOpt || 'Admin force logout (all)';
  list.forEach(u => { try { adminForceLogout(u.analyst_id, msg); } catch(e){} });
  return { ok:true, count:list.length };
}
