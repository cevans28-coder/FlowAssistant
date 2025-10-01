/******************************************************
 * 06_calendar.gs — Calendar sync & minutes helpers (optimized)
 * Depends on:
 * - TZ, SHEETS (00_constants.gs)
 * - toISODate_, normId_, indexMap_, readRows_(), deleteRowsBy_() (01/02 utils)
 * - master_(), getOrCreateMasterSheet_() (02_master_access.gs)
 * - getCurrentAnalystId_(), requireSession_() (03_sessions.gs)
 * - refreshLiveFor_() (04_state_engine.gs)
 ******************************************************/

/** ----------------------------------------------------
 * SHEET GUARDIAN — uses SHEETS.CAL_PULL (CalendarPull_v2)
 * --------------------------------------------------*/
function getCalPullSheet_() {
  const HEADERS = [
    'date', // yyyy-MM-dd (TZ)
    'analyst_id', // normalized email
    'start_iso', // event start ISO
    'end_iso', // event end ISO
    'title', // event title
    'category', // Meeting/Training/Coaching/Admin/Lunch (mapped)
    'my_status', // accepted | tentative | yes (legacy 'YES' ok)
    'source_id' // Calendar event id (stable; used for upsert key)
  ];
  const sh = getOrCreateMasterSheet_(SHEETS.CAL_PULL, HEADERS);

  // Realign first row if headers drifted
  const current = sh.getRange(1, 1, 1, Math.max(HEADERS.length, sh.getLastColumn()))
                    .getValues()[0].map(String);
  HEADERS.forEach((h, i) => { if ((current[i] || '').trim() !== h) sh.getRange(1, i + 1).setValue(h); });
  sh.setFrozenRows(1);
  return sh;
}

/** Load keyword→category mappings from Config (unchanged) */
function loadConfigMappings_() {
  const sh = master_().getSheetByName(SHEETS.CONFIG);
  if (!sh || sh.getLastRow() <= 1) return [];
  const values = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  return values.map(r => ({
    keyword: String(r[0]).toLowerCase().trim(),
    category: (String(r[1] || '').trim() || 'Meeting')
  }));
}

/** Map title → category (unchanged) */
function mapCategory_(title, mappings) {
  const t = String(title || '').toLowerCase();
  for (const m of (mappings || [])) {
    if (m.keyword && t.indexOf(m.keyword) !== -1) return m.category || 'Meeting';
  }
  return 'Meeting';
}

/** Date bounds helper (unchanged) */
function dayBounds_(dateISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) throw new Error('Use YYYY-MM-DD');
  const start = Utilities.parseDate(dateISO + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss');
  const end = Utilities.parseDate(dateISO + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');
  return { start, end };
}

/** Delete rows for this analyst+date (robust to Date vs string in the sheet) */
function _cal_deleteSlice_(sh, dateISO, analystIdNorm){
  if (sh.getLastRow() < 2) return;

  const tz = TZ || Session.getScriptTimeZone();
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);

  const cDate = idx['date'];
  const cAid = idx['analyst_id'];

  if (cDate == null || cAid == null) return;

  const toISO = (v) => {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const s = String(v || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    return '';
  };

  const rows = [];
  for (let r = 1; r < vals.length; r++) {
    const rowDateISO = toISO(vals[r][cDate]);
    const rowAnalyst = normId_(vals[r][cAid] || '');
    if (rowDateISO === dateISO && rowAnalyst === analystIdNorm) {
      rows.push(r + 1);
    }
  }
  for (let i = rows.length - 1; i >= 0; i--) sh.deleteRow(rows[i]);
}

/** ----------------------------------------------------
 * NEW: Upsert writer for CalendarPull_v2
 * - rows: Array<{date, analyst_id, start_iso, end_iso, title, category, my_status, source_id}>
 * - key = date|analyst_id_norm|source_id
 * --------------------------------------------------*/
function upsertCalendarPull(rows) {
  rows = Array.isArray(rows) ? rows : [];
  if (!rows.length) return { ok:true, updated:0, inserted:0, total:0 };

  const sh = getCalPullSheet_();
  const last = sh.getLastRow();
  const hdr = sh.getRange(1,1,1,Math.max(8, sh.getLastColumn())).getValues()[0].map(String);
  const idx = indexMap_(hdr);

  const cDate = idx['date'], cAid = idx['analyst_id'], cStart = idx['start_iso'], cEnd = idx['end_iso'];
  const cTitle = idx['title'], cCat = idx['category'], cMs = idx['my_status'], cSrc = idx['source_id'];

  // Build in-memory index of existing keys -> row number
  const map = {};
  if (last >= 2) {
    const existing = sh.getRange(2, 1, last-1, hdr.length).getValues();
    for (let i=0; i<existing.length; i++){
      const r = existing[i];
      const key = [
        String(r[cDate] || '').trim(),
        normId_(r[cAid] || ''),
        String(r[cSrc] || '').trim()
      ].join('|');
      map[key] = 2 + i; // 1-based row index
    }
  }

  const updates = []; // [rowNumber, [values...]]
  const inserts = [];

  rows.forEach(obj => {
    const date = String(obj.date || obj.date_iso || '').trim();
    const aid = normId_(obj.analyst_id || obj.email || '');
    const src = String(obj.source_id || obj.event_id || '').trim();
    if (!date || !aid || !src) return;

    const rec = [
      date,
      aid,
      String(obj.start_iso || ''),
      String(obj.end_iso || ''),
      String(obj.title || ''),
      String(obj.category || ''),
      String(obj.my_status || '').toLowerCase(), // normalize; OK if 'accepted'/'tentative'/'yes'
      src
    ];

    const key = [date, aid, src].join('|');
    const rowNum = map[key];
    if (rowNum) {
      updates.push([rowNum, rec]);
    } else {
      inserts.push(rec);
    }
  });

  // Batch apply updates
  if (updates.length) {
    // Group by contiguous blocks when possible (opt: keep simple per-row sets)
    updates.forEach(([rowNum, rec])=>{
      sh.getRange(rowNum, 1, 1, rec.length).setValues([rec]);
    });
  }

  // Append inserts
  if (inserts.length) {
    sh.getRange(sh.getLastRow()+1, 1, inserts.length, inserts[0].length).setValues(inserts);
  }

  return { ok:true, updated: updates.length, inserted: inserts.length, total: updates.length + inserts.length };
}

/** Heuristic: determine if an event is Out-Of-Office */
function isOOOTitle_(title) {
  const t = String(title || '').toLowerCase();
  // common phrases; tweak to your org’s terms
  return /\boo+(\b|$)|\bout of office\b|\bannual leave\b|\bholiday\b|\bvaca(tion)?\b|\bsick\b|\bpto\b/.test(t);
}

/**
 * Sync ONLY accepted calendar events for a specific date into CalendarPull_v2.
 * - Treat YES / OWNER / ORGANIZER as accepted.
 * - Skip personal events with 0 guests; include 1:1 (>=1 guest is ok).
 * - Fully rewrites this analyst+date slice (robust delete), with extra in-memory de-dupe.
 */
function syncMyCalendarForDate(token, dateISO) {
  if (token) token = requireSessionOrAdopt_(token);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) throw new Error('Use YYYY-MM-DD');

  const tz = TZ || Session.getScriptTimeZone();
  const idRaw = getCurrentAnalystId_();
  const idNorm = normId_(idRaw); // <- canonical id in the sheet
  const { start, end } = dayBounds_(dateISO);
  const sh = getCalPullSheet_();
  const mappings = loadConfigMappings_();

  // Helper: detect OOO via config or keywords
  const looksOOO = (title) => {
    const mapped = mapCategory_(title, mappings);
    if (String(mapped).toLowerCase() === 'ooo') return true;
    return /\bout of office\b|\booo\b|\bannual leave\b|\bholiday\b|\bvaca(tion)?\b|\bpto\b|\bsick\b/i.test(String(title||''));
  };

  // Hard-delete the existing slice (robust to Date vs string)
  _cal_deleteSlice_(sh, dateISO, idNorm);

  const cal = CalendarApp.getDefaultCalendar();
  const events = cal.getEvents(start, end) || [];

  let scanned = 0, acceptedKept = 0, oooKept = 0;
  const rows = [];

  for (const ev of events) {
    scanned++;

    // Times
    const s = ev.getStartTime(); const e = ev.getEndTime();
    if (!(s instanceof Date) || isNaN(s) || !(e instanceof Date) || isNaN(e) || e <= s) continue;

    // Core bits
    const title = ev.getTitle() || '';
    let myStatus = 'UNKNOWN';
    try { myStatus = String(ev.getMyStatus()).toUpperCase(); } catch (e) {}
    let guestCount = 0;
    try { guestCount = (ev.getGuestList() || []).length; } catch (e) { guestCount = 0; }

    // OOO: force tag + keep always
    if (looksOOO(title)) {
      rows.push([
        dateISO,
        idNorm, // <- write normalised id
        s.toISOString(),
        e.toISOString(),
        title,
        'OOO', // <- force category
        'YES', // <- store as accepted so totals logic is simple
        ev.getId()
      ]);
      oooKept++;
      continue;
    }

    // Non-OOO acceptance rule: YES/ACCEPTED/OWNER/ORGANIZER/TENTATIVE
    const isAccepted = ['YES','ACCEPTED','OWNER','ORGANIZER','TENTATIVE'].includes(myStatus);
    if (!isAccepted) continue;

    // Include 1:1s, exclude solo: require at least one guest
    if (guestCount < 1) continue;

    // Map category (default to Meeting)
    const mappedCategory = mapCategory_(title, mappings) || 'Meeting';

    rows.push([
      dateISO,
      idNorm, // <- write normalised id
      s.toISOString(),
      e.toISOString(),
      title,
      mappedCategory,
      'YES',
      ev.getId()
    ]);
    acceptedKept++;
  }

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Keep LIVE + Today summary fresh (best-effort)
  try { refreshLiveFor_(idNorm); } catch (e) {}
  try { _bustTodaySummaryCache_(idNorm); } catch (e) {}

  return {
    ok: true,
    date: dateISO,
    analyst_id: idNorm,
    scanned,
    kept_meetings: acceptedKept,
    kept_ooo: oooKept,
    written_rows: rows.length
  };
}

/** Today convenience */
function syncMyCalendarTodayNoToken() {
  const today = toISODate_(new Date());
  return syncMyCalendarForDate(null, today);
}

/**
 * Sum minutes of OOO calendar entries for an analyst on a date.
 * Looks at CalendarPull_v2 and totals rows with category='OOO' or OOO-looking titles.
 */
function getOOOMinutesForDay_(analystId, dateISO) {
  if (!analystId || !/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO||''))) return 0;

  const ss = master_();
  const tz = TZ || Session.getScriptTimeZone();

  const names = [];
  if (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.CAL_PULL) names.push(SHEETS.CAL_PULL);
  names.push('CalendarPull_v2', 'CalendarPull');

  let sh = null;
  for (const n of names) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) { sh = s; break; }
  }
  if (!sh) return 0;

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);

  const cDate = idx['date'];
  const cAid = idx['analyst_id'];
  const cTitle = idx['title'];
  const cCat = idx['category'];
  const cStart = idx['start_iso'];
  const cEnd = idx['end_iso'];

  if ([cDate,cAid,cCat,cStart,cEnd].some(x => x == null)) return 0;

  const wantId = normId_(analystId);

  const toISO = (v) => {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const s = String(v || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    return '';
  };

  let total = 0;
  const seen = new Set(); // de-dup by start|end|title

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (toISO(row[cDate]) !== dateISO) continue;
    if (normId_(row[cAid] || '') !== wantId) continue;

    const cat = String(row[cCat] || '').toLowerCase();
    const title = String(row[cTitle] || '');
    const oooLike = (cat === 'ooo') || /\bout of office\b|\booo\b|\bannual leave\b|\bholiday\b|\bvaca(tion)?\b|\bpto\b|\bsick\b/i.test(title);
    if (!oooLike) continue;

    const sIso = String(row[cStart] || '');
    const eIso = String(row[cEnd] || '');
    if (!sIso || !eIso) continue;

    const key = sIso + '|' + eIso + '|' + title;
    if (seen.has(key)) continue;
    seen.add(key);

    const ms = (new Date(eIso)) - (new Date(sIso));
    if (isFinite(ms) && ms > 0) total += Math.round(ms / 60000);
  }

  return Math.max(0, total);
}

/**
 * Sum minutes of ACCEPTED calendar entries for an analyst on a date.
 * Treat YES/OWNER/ORGANIZER as accepted. Guard against duplicate rows.
 */
function getAcceptedMeetingMinutes_(analystId, dateISO) {
  if (!analystId || !/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO || ''))) return 0;

  const ss = master_();
  const tz = TZ || Session.getScriptTimeZone();

  const names = [];
  if (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.CAL_PULL) names.push(SHEETS.CAL_PULL);
  names.push('CalendarPull_v2', 'CalendarPull');

  let sh = null;
  for (const n of names) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) { sh = s; break; }
  }
  if (!sh) return 0;

  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return 0;

  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);
  const cDate = idx['date'],
        cAid = idx['analyst_id'],
        cMs = idx['my_status'],
        cStart = idx['start_iso'],
        cEnd = idx['end_iso'],
        cSrc = idx['source_id'],
        cCat = idx['category']; // added so we can detect OOO rows

  if ([cDate, cAid, cMs, cStart, cEnd].some(x => x == null)) return 0;

  const wantId = normId_(analystId);

  const toISO = (v) => {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const s = String(v || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    return '';
  };

  let total = 0;
  const seen = new Set(); // de-dup by source_id|start_iso|end_iso

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (toISO(row[cDate]) !== dateISO) continue;
    if (normId_(row[cAid] || '') !== wantId) continue;

    // Skip OOO rows entirely
    if (cCat != null && String(row[cCat] || '').trim().toLowerCase() === 'ooo') continue;

    const status = String(row[cMs] || '').toUpperCase();
    if (!['YES','OWNER','ORGANIZER','ACCEPTED'].includes(status)) continue;

    const sIso = String(row[cStart] || '');
    const eIso = String(row[cEnd] || '');
    if (!sIso || !eIso) continue;

    const key = String(row[cSrc] || '') + '|' + sIso + '|' + eIso;
    if (seen.has(key)) continue;
    seen.add(key);

    const ms = (new Date(eIso)) - (new Date(sIso));
    if (isFinite(ms) && ms > 0) total += Math.round(ms / 60000);
  }

  return Math.max(0, total);
}

function CAL_Diag_Today(){
  const id = getCurrentAnalystId_();
  const day = toISODate_(new Date());
  const sh = master_().getSheetByName(SHEETS.CAL_PULL);
  if (!sh) { Logger.log('No CAL_PULL sheet.'); return; }
  const rows = readRows_(sh).filter(r => r.date_str === day && r.analyst_id_norm === id);
  Logger.log('Rows today: %s', rows.length);
  rows.forEach(r => Logger.log('%s .. %s | status=%s | title=%s',
    r.start_iso, r.end_iso, r.my_status, r.title));

  const mins = getAcceptedMeetingMinutes_(id, day);
  Logger.log('getAcceptedMeetingMinutes_ => %s minutes', mins);
}

function CAL_Diag_TodaySummary(){
  _bustTodaySummaryCache_(getCurrentAnalystId_());
  Logger.log(JSON.stringify(getTodaySummary()));
}

function CAL_Diag_TodayMinutes(){
  const id = getCurrentAnalystId_();
  const day = toISODate_(new Date());
  Logger.log('Accepted meeting mins today: %s', getAcceptedMeetingMinutes_(id, day));
}

function CAL_Dump_TodayRows(){
  const id = getCurrentAnalystId_();
  const day = toISODate_(new Date());
  const tz = TZ || Session.getScriptTimeZone();

  const ss = master_();
  const names = [];
  if (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.CAL_PULL) names.push(SHEETS.CAL_PULL);
  names.push('CalendarPull_v2', 'CalendarPull');

  let sh = null;
  for (const n of names) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) { sh = s; break; }
  }
  if (!sh) { 
    Logger.log('No CalendarPull sheet found. Tried: %s', JSON.stringify(names)); 
    return; 
  }

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);

  Logger.log('Using sheet: %s | Headers: %s', sh.getName(), JSON.stringify(hdr));

  const cDate = idx['date'];
  const cAid = idx['analyst_id'];
  const cMs = idx['my_status'];
  const cStart = idx['start_iso'];
  const cEnd = idx['end_iso'];
  const cCat = idx['category'];

  const toISO = (v) => {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const s = String(v || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    return '';
  };

  const want = normId_(id);
  let total = 0;
  const rows = [];

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const rowDay = toISO(row[cDate]);
    const aid = normId_(row[cAid] || '');
    if (rowDay !== day || aid !== want) continue;

    const status = String(row[cMs] || '');
    const sIso = String(row[cStart] || '');
    const eIso = String(row[cEnd] || '');
    const mins = (sIso && eIso) ? Math.max(0, Math.round((new Date(eIso) - new Date(sIso))/60000)) : 0;

    // Skip OOO rows
    const cat = cCat != null ? String(row[cCat] || '').toLowerCase() : '';
    const isOOO = (cat === 'ooo' || isOOOTitle_(String(row[idx['title']] || '')));

    rows.push({ rowDay, status, category: cat, start_iso: sIso, end_iso: eIso, mins, skippedOOO: isOOO });

    if (!isOOO && ['accepted','Accept','ACCEPTED','YES','Yes','yes'].includes(status)) {
      total += mins;
    }
  }

  Logger.log('Rows for me today: %s', JSON.stringify(rows, null, 2));
  Logger.log('Total accepted mins (diag, excl OOO): %s', total);
}

/** True if CalendarPull_v2 has any OOO entry for (analystId, dateISO) */
function isOOODayInPull_(analystId, dateISO) {
  if (!analystId || !/^\d{4}-\d{2}-\d{2}$/.test(String(dateISO||''))) return false;

  const ss = master_();
  const names = [];
  if (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.CAL_PULL) names.push(SHEETS.CAL_PULL);
  names.push('CalendarPull_v2', 'CalendarPull');

  let sh = null;
  for (const n of names) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) { sh = s; break; }
  }
  if (!sh) return false;

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const idx = indexMap_(hdr);

  const cDate = idx['date'];
  const cAid = idx['analyst_id'];
  const cCat = idx['category'];

  if (cDate == null || cAid == null || cCat == null) return false;
  const want = normId_(analystId);

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const d = toISODate_(row[cDate]); // reuse your own robust to-ISO helper if you have one
    if (d !== dateISO) continue;
    if (normId_(row[cAid]) !== want) continue;
    const cat = String(row[cCat] || '').toLowerCase();
    if (cat === 'ooo' || isOOOTitle_(String(row[idx['title']] || ''))) return true;
  }
  return false;
}

function CAL_Diag_ListToday() {
  const tz = Session.getScriptTimeZone();
  const day = Utilities.formatDate(new Date(), tz, '2025-09-11');
  const evs = CalendarApp.getDefaultCalendar().getEvents(
    Utilities.parseDate(day+' 00:00:00', tz, 'yyyy-MM-dd HH:mm:ss'),
    Utilities.parseDate(day+' 23:59:59', tz, 'yyyy-MM-dd HH:mm:ss')
  );
  const mappings = loadConfigMappings_();
  evs.forEach(ev=>{
    const title = ev.getTitle() || '';
    const cat = mapCategory_(title, mappings);
    const isOOO = (String(cat).toLowerCase()==='ooo') ||
                  /\bout of office\b|\booo\b|\bannual leave\b|\bholiday\b|\bpto\b|\bsick\b/i.test(title);
    Logger.log('%s | cat=%s | OOO=%s | guests=%s | status=%s',
      title, cat, isOOO, (ev.getGuestList()||[]).length, ev.getMyStatus());
  });
}
