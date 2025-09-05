/******************************************************
 * 06_calendar.gs — Calendar sync & minutes helpers
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_, dayBounds_ (01_utils.gs)
 * - master_(), getOrCreateMasterSheet_(), readRows_(), deleteRowsBy_(), indexMap_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_() (03_sessions.gs)
 * - refreshLiveFor_() (04_state_engine.gs)
 ******************************************************/

/**
 * Load keyword→category mappings from Config.
 * Rows: keyword | category
 * - keyword is matched case-insensitively within the event title
 * - blank/unknown category defaults to 'Meeting'
 */
function loadConfigMappings_() {
  const sh = master_().getSheetByName(SHEETS.CONFIG);
  if (!sh || sh.getLastRow() <= 1) return [];
  const values = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  return values.map(r => ({
    keyword: String(r[0]).toLowerCase().trim(),
    category: (String(r[1] || '').trim() || 'Meeting')
  }));
}

/**
 * Given an event title and mappings, return a category or '' to skip.
 * - If a mapping keyword is found inside the title, returns its category (or 'Meeting')
 * - If no mapping hits, return '' to SKIP (so only configured items are imported)
 * (If you want unmatched to fall back to 'Meeting', change the last line.)
 */
function mapCategory_(title, mappings) {
  const t = String(title || '').toLowerCase();
  for (const m of mappings) {
    if (m.keyword && t.indexOf(m.keyword) !== -1) return m.category || 'Meeting';
  }
  return ''; // unmatched → skip
  // return 'Meeting'; // <- alternative: default to 'Meeting' if you prefer importing all
}

/**
 * Sync ONLY your accepted calendar events for a specific date into CAL_PULL.
 * - Filters to events where getMyStatus() is YES (accepted)
 * - Maps category using Config sheet keywords
 * - Clears existing rows for that date+analyst before writing
 * - Does NOT double-count; uses a full rewrite for the date
 */
function syncMyCalendarForDate(token, dateISO) {
  requireSession_(token);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || '')))
    throw new Error('Use YYYY-MM-DD');

  const id = getCurrentAnalystId_();
  const { start, end } = dayBounds_(dateISO);
  const events = CalendarApp.getDefaultCalendar().getEvents(start, end) || [];
  const mappings = loadConfigMappings_();

  const sh = getOrCreateMasterSheet_(SHEETS.CAL_PULL, [
    'date','analyst_id','start_iso','end_iso','title','category','my_status','source_id'
  ]);

  // Remove existing rows for this date+analyst (fresh rebuild model)
  deleteRowsBy_(sh, row => String(row[0]) === dateISO && normId_(row[1]) === id);

  const rows = [];
  for (const ev of events) {
    // Only accepted events
    let accepted = false;
    try {
      accepted = ev.getMyStatus() === CalendarApp.GuestStatus.YES;
    } catch (e) {
      accepted = false;
    }
    if (!accepted) continue;

    const title = ev.getTitle() || '';
    const category = mapCategory_(title, mappings);
    if (!category) continue; // skip events that don't match your keywords

    rows.push([
      dateISO,
      id,
      ev.getStartTime().toISOString(),
      ev.getEndTime().toISOString(),
      title,
      category,
      'Accepted', // normalised value
      ev.getId()
    ]);
  }

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Keep LIVE fresh (useful for nudges/badges)
  try { refreshLiveFor_(id); } catch (e) {}
try { updateLiveKPIsFor_(id); } catch(e) {}

  return { ok: true, date: dateISO, analyst_id: id, added: rows.length };
}

/** Convenience: today’s sync without a token (menu item). */
function syncMyCalendarTodayNoToken() {
  return syncMyCalendarForDate(null, toISODate_(new Date()));
}

/**
 * Sum total minutes of ACCEPTED calendar entries for an analyst on a date.
 * Reads from CAL_PULL (not live Calendar) for speed and consistency.
 * Recognises my_status values like 'Accepted', 'YES', 'Yes', 'yes'.
 */
function getAcceptedMeetingMinutes_(analystId, dateISO) {
  const ss = master_();
  const sh = ss.getSheetByName(SHEETS.CAL_PULL);
  if (!sh || sh.getLastRow() <= 1) return 0;

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);

  const wantId = normId_(analystId);
  let total = 0;

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];

    if (String(row[idx['date']]) !== dateISO) continue;
    if (normId_(row[idx['analyst_id']]) !== wantId) continue;

    const status = String(row[idx['my_status']] || '').toLowerCase();
    if (!['accepted', 'accept', 'yes'].includes(status)) continue;

    const startIso = String(row[idx['start_iso']] || '');
    const endIso = String(row[idx['end_iso']] || '');
    if (!startIso || !endIso) continue;

    const mins = Math.max(0, (new Date(endIso) - new Date(startIso)) / 60000);
    total += Math.round(mins);
  }
  return total;
}

function dayBounds_(dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO||''))) {
    throw new Error('Use YYYY-MM-DD');
  }
  const start = Utilities.parseDate(dateISO + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss');
  const end = Utilities.parseDate(dateISO + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');
  return { start, end };
}
