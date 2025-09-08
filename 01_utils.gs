/******************************************************
 * 01_utils.gs — Core utilities (refined)
 * - Safe, tiny helpers used across the project
 * - Careful not to change existing behaviour/signatures
 ******************************************************/

/**
 * Build a { headerName: columnIndex } map.
 * Assumes headers are unique in the row.
 */
function indexMap_(headers) {
  const m = Object.create(null);
  (headers || []).forEach((h, i) => { m[String(h).trim()] = i; });
  return m;
}

/** Normalise analyst IDs/emails consistently (lowercase+trim). */
function normId_(s) {
  return String(s || '').toLowerCase().trim();
}

/** Format Date → 'YYYY-MM-DD' in project TZ (no time). */
function toISODate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
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
  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length <= 1) return [];

  const hdr = vals[0].map(h => String(h).trim());
  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    // Fast skip for fully empty rows
    if (!row || row.every(c => c === '' || c === null)) continue;

    const o = {};
    // Map by header name
    for (let i = 0; i < hdr.length; i++) o[hdr[i]] = row[i];

    // Normalised analyst id helper
    o.analyst_id_norm = normId_(o['analyst_id']);

    // Parse canonical timestamp fields
    const tsIso = o['timestamp_iso'] || o['completed_at_iso'] || null;
    if (tsIso) {
      const ts = new Date(tsIso);
      if (!isNaN(ts)) o.ts = ts;
    }
    // Optional range timestamps
    if (o['start_iso']) { const s = new Date(o['start_iso']); if (!isNaN(s)) o.start = s; }
    if (o['end_iso']) { const e = new Date(o['end_iso']); if (!isNaN(e)) o.end = e; }

    // Derive normalised 'YYYY-MM-DD' date string
    let dateStr = '';
    if (o.ts instanceof Date && !isNaN(o.ts)) {
      dateStr = Utilities.formatDate(o.ts, TZ, 'yyyy-MM-dd');
    } else {
      const raw = o['date'];
      if (raw instanceof Date && !isNaN(raw)) {
        dateStr = Utilities.formatDate(raw, TZ, 'yyyy-MM-dd');
      } else if (typeof raw === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(raw.trim())) {
        dateStr = raw.trim();
      } else {
        // last resort: keep whatever was there (for debugging)
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
  const v = sh.getDataRange().getValues();
  if (!v || v.length <= 1) return;

  const hdr = v[0];
  const keep = v.slice(1).filter(r => !pred(r));
  sh.clearContents();
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  if (keep.length) sh.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
}

/* ===========================================================
 * Shared date helpers (centralised here for reuse)
 * =========================================================== */

/**
 * Normalise many date inputs to 'YYYY-MM-DD' in TZ.
 * Accepts:
 * - 'YYYY-MM-DD' (returns as-is)
 * - 'DD/MM/YYYY' and 'MM/DD/YYYY' (best-effort, UK-first)
 * - Date objects
 * - ISO timestamps 'YYYY-MM-DDTHH:mm:ssZ'
 * - Other parsable strings (fallback via new Date())
 *
 * This mirrors the tolerant version you used in 09_tl_api.
 * Keep it here so other files can call the same function.
 */
function normaliseToISODate_(input, tz) {
  const _tz = tz || TZ;
  if (!input) return Utilities.formatDate(new Date(), _tz, 'yyyy-MM-dd');

  // Date object
  if (input instanceof Date && !isNaN(input)) {
    return Utilities.formatDate(input, _tz, 'yyyy-MM-dd');
  }

  const s = String(input).trim();

  // Already ISO YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // ISO timestamp → Date → ISO
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, _tz, 'yyyy-MM-dd');
  }

  // DD/MM/YYYY or MM/DD/YYYY
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    let d = parseInt(m[1], 10);
    let mo = parseInt(m[2], 10);
    const y = parseInt(m[3], 10);

    // Disambiguate generously (prefer UK style when ambiguous)
    if (d > 12) {
      // DD/MM — fine
    } else if (mo > 12) {
      // Impossible month -> must be DD/MM (swap)
      const tmp = d; d = mo; mo = tmp;
    } else {
      // Ambiguous like 04/05/2025 → assume DD/MM (UK style)
      // (no swap: we already treat first as day)
    }

    const jsDate = new Date(y, mo - 1, d, 12, 0, 0); // Noon avoids DST edges
    return Utilities.formatDate(jsDate, _tz, 'yyyy-MM-dd');
  }

  // Fallback to Date()
  const tryDate = new Date(s);
  if (!isNaN(tryDate)) {
    return Utilities.formatDate(tryDate, _tz, 'yyyy-MM-dd');
  }

  // Last resort: today
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
  const d = String(dateISO || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Use YYYY-MM-DD');

  const start = Utilities.parseDate(d + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss');
  let end = Utilities.parseDate(d + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');

  if (capToNow !== false) {
    const todayISO = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (d === todayISO) {
      const now = new Date();
      if (now < end) end = now;
    }
  }

  if (end <= start) end = new Date(start.getTime() + 60 * 1000); // guard: 1 min window
  return { start, end };
}

function adminRecountTodayChecksAll() {
  const ss = master_();
  const live = ss.getSheetByName(SHEETS.LIVE);
  const today = toISODate_(new Date());
  if (!live || live.getLastRow() < 2) return { ok:false, updated:0 };

  const vals = live.getDataRange().getValues();
  const hdr = vals[0].map(String);
  const L = indexMap_(hdr);

  let n=0;
  for (let r=1; r<vals.length; r++) {
    const id = String(vals[r][L['analyst_id']]||'').trim();
    if (!id) continue;
    const count = readRows_(ss.getSheetByName(SHEETS.CHECK_EVENTS))
      .filter(x => x.date_str===today && x.analyst_id_norm===normId_(id)).length;
    if (L['today_checks'] != null) {
      live.getRange(r+1, L['today_checks']+1).setValue(count);
      n++;
    }
  }
  return { ok:true, updated:n };
}


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
    // If your Exceptions module exposes a suitable method:
    if (typeof Exceptions !== 'undefined') {
      // Try a couple of common shapes
      if (typeof Exceptions.listMineForDate === 'function') {
        return Exceptions.listMineForDate(dateISO); // expected: array of items
      }
      if (typeof Exceptions.listForAnalystDate === 'function') {
        return Exceptions.listForAnalystDate(me, dateISO); // alternate signature
      }
    }
  } catch (e) {
    // Swallow and fall through to empty set
  }
  return []; // no exceptions or module not present
}

function Exceptions_listMineForDate(dateISO) {
  return getMyExceptionsForDate(dateISO);
}
