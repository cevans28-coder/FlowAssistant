/******************************************************
 * 06_calendar.gs — Calendar sync & minutes helpers (optimized)
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_, indexMap_, readRows_(), deleteRowsBy_() (01/02 utils)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_() (03_sessions.gs)
 * - refreshLiveFor_() (04_state_engine.gs)
 ******************************************************/

/* ----------------------------------------------------
 * SHEET GUARDIANS
 * --------------------------------------------------*/
/**
 * Ensure and return the canonical CalendarPull sheet with fixed headers.
 * Prevents column drift and lets other code assume stable indices.
 */
function getCalPullSheet_() {
  const HEADERS = [
    'date', // yyyy-MM-dd (TZ)
    'analyst_id', // normalized email
    'start_iso', // event start ISO
    'end_iso', // event end ISO
    'title', // event title
    'category', // Meeting/Training/Coaching/Admin/Lunch (mapped)
    'my_status', // YES (accepted)
    'source_id' // Calendar event id (for de-dupe/debug)
  ];
  const sh = getOrCreateMasterSheet_(SHEETS.CAL_PULL, HEADERS);

  // Realign first row if headers drifted
  const current = sh.getRange(1, 1, 1, Math.max(HEADERS.length, sh.getLastColumn()))
                    .getValues()[0].map(String);
  HEADERS.forEach((h, i) => { if ((current[i] || '').trim() !== h) sh.getRange(1, i + 1).setValue(h); });
  sh.setFrozenRows(1);
  return sh;
}

/* ----------------------------------------------------
 * CONFIG → CATEGORY MAPPING
 * --------------------------------------------------*/
/**
 * Load keyword→category mappings from Config.
 * Rows: keyword | category
 * - Case-insensitive contains() match on title
 * - Missing category defaults to 'Meeting'
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
 * Given an event title and mappings, return a category.
 * - First keyword hit wins
 * - If no mapping hits, we still include the event as 'Meeting'
 */
function mapCategory_(title, mappings) {
  const t = String(title || '').toLowerCase();
  for (const m of (mappings || [])) {
    if (m.keyword && t.indexOf(m.keyword) !== -1) return m.category || 'Meeting';
  }
  return 'Meeting';
}

/* ----------------------------------------------------
 * DATE BOUNDS
 * --------------------------------------------------*/
/** Start/end Date objects for a yyyy-MM-dd in TZ. */
function dayBounds_(dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) throw new Error('Use YYYY-MM-DD');
  const start = Utilities.parseDate(dateISO + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss');
  const end = Utilities.parseDate(dateISO + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');
  return { start, end };
}

/* ----------------------------------------------------
 * CALENDAR → MASTER PULL
 * --------------------------------------------------*/
/**
 * Sync ONLY accepted calendar events for a specific date into CAL_PULL.
 * Filters:
 * - getMyStatus() must be YES (accepted)
 * - Event must have >1 guest (i.e., a multi-attendee meeting)
 * Mapping:
 * - Title is mapped via Config keywords to a category (default 'Meeting')
 * Write pattern:
 * - Fully rewrites this analyst+date slice in CAL_PULL (deletes old rows, appends fresh)
 */
function syncMyCalendarForDate(token, dateISO) {
  // Token is required by your UI, but we accept null for menu/trigger based usage
  if (token) requireSession_(token);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) throw new Error('Use YYYY-MM-DD');

  const id = getCurrentAnalystId_();
  const { start, end } = dayBounds_(dateISO);

  // Pull events in one go
  const events = CalendarApp.getDefaultCalendar().getEvents(start, end) || [];
  const mappings = loadConfigMappings_();

  // Prepare target sheet & clear previous slice
  const sh = getCalPullSheet_();
  deleteRowsBy_(sh, row => String(row[0]) === dateISO && normId_(row[1]) === id);

  // Build rows (only accepted + >1 guest)
  const rows = [];
  for (const ev of events) {
    // Must be accepted by me
    let accepted = false;
    try { accepted = ev.getMyStatus() === CalendarApp.GuestStatus.YES; } catch (e) {}
    if (!accepted) continue;

    // Must be multi-attendee meeting (>1 guest entries)
    let guestCount = 0;
    try { guestCount = (ev.getGuestList() || []).length; } catch (e) { guestCount = 0; }
    if (guestCount <= 1) continue;

    const title = ev.getTitle() || '';
    const category = mapCategory_(title, mappings); // always returns a non-empty string

    rows.push([
      dateISO,
      id,
      ev.getStartTime().toISOString(),
      ev.getEndTime().toISOString(),
      title,
      category,
      'YES',
      ev.getId()
    ]);
  }

  // Bulk-append
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Keep LIVE fresh for this user
  try { refreshLiveFor_(id); } catch (e) {}

  return { ok: true, date: dateISO, analyst_id: id, added: rows.length };
}

/** Convenience: today’s sync without a token (menu item / trigger). */
function syncMyCalendarTodayNoToken() {
  return syncMyCalendarForDate(null, toISODate_(new Date()));
}

/* ----------------------------------------------------
 * READBACK HELPERS
 * --------------------------------------------------*/
/**
 * Sum total minutes of ACCEPTED calendar entries for an analyst on a date.
 * Reads from CAL_PULL (not live Calendar) for speed/consistency.
 * Recognises my_status values like 'YES' or common 'Accepted' variants.
 */
function getAcceptedMeetingMinutes_(analystId, dateISO) {
  const key = 'MEET_MINS:' + normId_(analystId) + ':' + dateISO;
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return Number(hit) || 0;

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
    if (!['accepted','accept','yes'].includes(status)) continue;

    const startIso = String(row[idx['start_iso']] || '');
    const endIso = String(row[idx['end_iso']] || '');
    if (!startIso || !endIso) continue;

    const mins = Math.max(0, (new Date(endIso) - new Date(startIso)) / 60000);
    total += Math.round(mins);
  }

  try { cache.put(key, String(total), 300); } catch(e) {} // 5 min
  return total;
}
