/******************************************************
 * 01_utils.gs — Core utilities (refined + v2 shims)
 * - Safe, tiny helpers used across the project
 * - Backward compatible with your existing code
 * - Adds helpers required by v2 materializers
 ******************************************************
 * Build a { headerName: columnIndex } map (exact, case-sensitive).
 * Assumes headers are unique in the row.
 */
function indexMap_(headers) {
  var m = Object.create(null);
  (headers || []).forEach(function(h, i){ m[String(h).trim()] = i; });
  return m;
}

/**
 * Loose index map:
 * - lowercases
 * - underscores spaces
 * - strips non [a-z0-9_]
 * Example: "Completed At (ISO)" -> "completed_at_iso"
 */
function indexMapLoose_(headers){
  var m = Object.create(null);
  (headers || []).forEach(function(h, i){
    var k = String(h||'').toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'');
    m[k] = i;
  });
  return m;
}

/* =========================
 * ID / Numbers / Dates
 * ========================= */

/** Normalise analyst IDs/emails consistently (lowercase+trim). */
function normId_(v) {
  const s = String(v || '').trim().toLowerCase();
  if (!s) return '';
  return s; // keep full email if present; do not chop after '@'
}
/** Small numeric normaliser: Number(x) if finite, else 0. */
function n_(x){ x = Number(x); return (Number.isFinite ? Number.isFinite(x) : isFinite(x)) ? x : 0; }

/** Format Date → 'YYYY-MM-DD' in project TZ (no time). */
function toISODate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/** Guarded ISO date checker 'YYYY-MM-DD'. */
function isISO_(d){ return /^\d{4}-\d{2}-\d{2}$/.test(String(d||'')); }

/**
 * Tolerant cell→ISO converter:
 * - Accepts Date objects, ISO strings, other parseable forms
 * - Returns '' on failure
 */
function asISO_(cell){
  if (cell instanceof Date && !isNaN(cell)) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(cell||'').trim();
  if (isISO_(s)) return s;
  var d = new Date(s);
  if (d instanceof Date && !isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return '';
}

/**
 * Whole minutes between two Date objects (non-negative, rounded).
 * Preserves your original rounding behaviour (Math.round).
 */
function minutesBetween_(a, b) {
  return Math.max(0, Math.round((b - a) / 60000));
}

/** Small rounding helpers used for KPI displays. */
function round0(n) { return Math.round(Number(n) || 0); }
function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
/** Keep 4 decimals on a *fraction* then convert to percent elsewhere if needed. */
function roundPct(n) { return Math.round((Number(n) || 0) * 10000) / 10000; }

/**
 * Normalise many date inputs to 'YYYY-MM-DD' in TZ.
 * Accepts Date, ISO date, ISO timestamp, DD/MM/YYYY (UK-first) and best-effort strings.
 */
function normaliseToISODate_(input, tz) {
  var _tz = tz || TZ;
  if (!input) return Utilities.formatDate(new Date(), _tz, 'yyyy-MM-dd');

  if (input instanceof Date && !isNaN(input)) {
    return Utilities.formatDate(input, _tz, 'yyyy-MM-dd');
  }

  var s = String(input).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    var d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, _tz, 'yyyy-MM-dd');
  }

  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var d2 = parseInt(m[1], 10);
    var mo = parseInt(m[2], 10);
    var y = parseInt(m[3], 10);
    if (d2 > 12) {
      // DD/MM — fine
    } else if (mo > 12) {
      // Impossible month -> DD/MM (swap)
      var tmp = d2; d2 = mo; mo = tmp;
    } // else ambiguous; keep UK-first
    var jsDate = new Date(y, mo - 1, d2, 12, 0, 0); // Noon avoids DST edges
    return Utilities.formatDate(jsDate, _tz, 'yyyy-MM-dd');
  }

  var tryDate = new Date(s);
  if (!isNaN(tryDate)) {
    return Utilities.formatDate(tryDate, _tz, 'yyyy-MM-dd');
  }

  return Utilities.formatDate(new Date(), _tz, 'yyyy-MM-dd');
}

/**
 * Compute start/end Date objects (in TZ) for a given ISO date.
 * Optionally caps end to "now" when that date is today.
 * @param {string} dateISO 'YYYY-MM-DD'
 * @param {boolean} capToNow default true
 * @returns {{start: Date, end: Date}}
 */
function computeDayBounds_(dateISO, capToNow) {
  var d = String(dateISO || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Use YYYY-MM-DD');

  var start = Utilities.parseDate(d + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss');
  var end = Utilities.parseDate(d + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');

  if (capToNow !== false) {
    var todayISO = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (d === todayISO) {
      var now = new Date();
      if (now < end) end = now;
    }
  }

  if (end <= start) end = new Date(start.getTime() + 60 * 1000); // guard: 1 min window
  return { start: start, end: end };
}

/* =========================
 * Sheet read/write helpers
 * ========================= */

/**
 * Read a sheet into an array of row objects with conveniences:
 * - Keys by header names
 * - `analyst_id_norm` (normalised)
 * - `ts` (Date from 'timestamp_iso' or 'completed_at_iso' if present)
 * - `start` / `end` (Date from 'start_iso' / 'end_iso' if present)
 * - `date_str` ('YYYY-MM-DD' from ts or 'date' column)
 *
 * NOTE: Skips fully empty rows to save downstream work.
 */
function readRows_(sh) {
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (!vals || vals.length <= 1) return [];

  var hdr = vals[0].map(function(h){ return String(h).trim(); });
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    if (!row || row.every(function(c){ return c === '' || c === null; })) continue;

    var o = {};
    for (var i = 0; i < hdr.length; i++) o[hdr[i]] = row[i];

    o.analyst_id_norm = normId_(o['analyst_id']);

    var tsIso = o['timestamp_iso'] || o['completed_at_iso'] || null;
    if (tsIso) {
      var ts = new Date(tsIso);
      if (!isNaN(ts)) o.ts = ts;
    }
    if (o['start_iso']) { var s = new Date(o['start_iso']); if (!isNaN(s)) o.start = s; }
    if (o['end_iso']) { var e = new Date(o['end_iso']); if (!isNaN(e)) o.end = e; }

    var dateStr = '';
    if (o.ts instanceof Date && !isNaN(o.ts)) {
      dateStr = Utilities.formatDate(o.ts, TZ, 'yyyy-MM-dd');
    } else {
      var raw = o['date'];
      if (raw instanceof Date && !isNaN(raw)) {
        dateStr = Utilities.formatDate(raw, TZ, 'yyyy-MM-dd');
      } else if (typeof raw === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(raw.trim())) {
        dateStr = raw.trim();
      } else {
        dateStr = String(raw || '').trim();
      }
    }
    o.date_str = dateStr;

    out.push(o);
  }
  return out;
}

/**
 * Delete rows matching predicate and keep header row intact.
 * Rewrites the sheet in one go to reduce API calls.
 * @param {Sheet} sh
 * @param {(row:Array)=>boolean} pred - return true to delete that data row
 */
function deleteRowsBy_(sh, pred) {
  if (!sh) return;
  var v = sh.getDataRange().getValues();
  if (!v || v.length <= 1) return;

  var hdr = v[0];
  var keep = v.slice(1).filter(function(r){ return !pred(r); });
  sh.clearContents();
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  if (keep.length) sh.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
}

/* =========================
 * Admin helper (optional)
 * ========================= */

function adminRecountTodayChecksAll() {
  try{
    var ss = master_();
    var live = ss.getSheetByName(SHEETS && SHEETS.LIVE ? SHEETS.LIVE : 'Live');
    var today = toISODate_(new Date());
    if (!live || live.getLastRow() < 2) return { ok:false, updated:0 };

    var vals = live.getDataRange().getValues();
    var hdr = vals[0].map(String);
    var L = indexMap_(hdr);

    var shChecks = ss.getSheetByName(SHEETS && SHEETS.CHECK_EVENTS ? SHEETS.CHECK_EVENTS : 'CheckEvents');
    if (!shChecks) return { ok:false, updated:0 };

    var allChecks = readRows_(shChecks);
    var updated = 0;

    for (var r=1; r<vals.length; r++) {
      var id = String(vals[r][L['analyst_id']]||'').trim();
      if (!id) continue;
      var count = allChecks.filter(function(x){ return x.date_str===today && x.analyst_id_norm===normId_(id); }).length;
      if (L['today_checks'] != null) {
        live.getRange(r+1, L['today_checks']+1).setValue(count);
        updated++;
      }
    }
    return { ok:true, updated:updated };
  } catch(e){
    return { ok:false, error:String(e) };
  }
}

/* =========================
 * Exceptions passthrough (safe)
 * ========================= */

/** Safe wrapper for Flow Assistant “Quick Glance”
 * Returns today's exceptions for the signed-in analyst.
 * If Exceptions module isn’t present, returns [] (no errors).
 */
function getMyExceptionsForDate(dateISO) {
  try {
    var me = (Session.getActiveUser().getEmail() || '').toLowerCase();
    if (!dateISO) {
      dateISO = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (typeof Exceptions !== 'undefined') {
      if (typeof Exceptions.listMineForDate === 'function') {
        return Exceptions.listMineForDate(dateISO);
      }
      if (typeof Exceptions.listForAnalystDate === 'function') {
        return Exceptions.listForAnalystDate(me, dateISO);
      }
    }
  } catch (e) {
    // swallow + fall through
  }
  return [];
}

function Exceptions_listMineForDate(dateISO) {
  return getMyExceptionsForDate(dateISO);
}
