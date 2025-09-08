/******************************************************
 * 02_master_access.gs — Smart Master resolver (consolidated)
 * Order of precedence to resolve the Master workbook:
 * 1) Global constant MASTER_SPREADSHEET_ID (00_constants.gs)
 * 2) Active spreadsheet (bound script)
 * 3) Script Property: MASTER_SPREADSHEET_ID
 * 4) Settings tab key: MASTER_ID
 * Provides:
 * - master_() — cached opener of the Master spreadsheet
 * - getOrCreateMasterSheet_(name, headers) — ensure tab + headers
 * - ensureSheet_(ss, name, headers) — internal helper
 ******************************************************/

// Tiny in-memory cache for this execution
var __MASTER_ID_CACHE = null;
var __MASTER_SS_CACHE = null;

/** Try to read a usable Master ID from multiple places (memoized). */
function getMasterId_() {
  if (__MASTER_ID_CACHE) return __MASTER_ID_CACHE;

  // 1) Global constant
  try {
    if (typeof MASTER_SPREADSHEET_ID !== 'undefined') {
      const v = String(MASTER_SPREADSHEET_ID || '').trim();
      if (v && v !== 'PASTE_YOUR_MASTER_SHEET_ID_HERE') {
        __MASTER_ID_CACHE = v;
        return v;
      }
    }
  } catch (e) { /* ignore */ }

  // 2) Active (bound) spreadsheet
  try {
    const as = SpreadsheetApp.getActive();
    if (as) {
      __MASTER_ID_CACHE = as.getId();
      return __MASTER_ID_CACHE;
    }
  } catch (e) { /* ignore */ }

  // 3) Script Property
  try {
    const prop = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');
    if (prop && prop.trim()) {
      __MASTER_ID_CACHE = prop.trim();
      return __MASTER_ID_CACHE;
    }
  } catch (e) { /* ignore */ }

  // 4) Settings tab fallback (only works if active SS is available)
  try {
    const as = SpreadsheetApp.getActive();
    if (as) {
      const set = as.getSheetByName('Settings');
      if (set && set.getLastRow() > 1) {
        const v = set.getDataRange().getValues();
        for (let r = 1; r < v.length; r++) {
          const key = String(v[r][0] || '').trim().toUpperCase();
          if (key === 'MASTER_ID') {
            const id = String(v[r][1] || '').trim();
            if (id) {
              __MASTER_ID_CACHE = id;
              return __MASTER_ID_CACHE;
            }
          }
        }
      }
    }
  } catch (e) { /* ignore */ }

  return ''; // not found
}

/** Open the Master spreadsheet with robust fallbacks + clear error if missing. Cached. */
function master_() {
  if (__MASTER_SS_CACHE) return __MASTER_SS_CACHE;

  const id = getMasterId_();
  if (id) {
    try {
      __MASTER_SS_CACHE = SpreadsheetApp.openById(id);
      return __MASTER_SS_CACHE;
    } catch (e) {
      throw new Error('MASTER_SPREADSHEET_ID appears set but cannot be opened. Check sharing/ID.');
    }
  }

  // Nothing resolved: give concrete next steps
  throw new Error(
    'MASTER_SPREADSHEET_ID is not configured.\n\n' +
    'Fix it via one of the following:\n' +
    ' A) In 00_constants.gs, set const MASTER_SPREADSHEET_ID = "YOUR_SHEET_ID";\n' +
    ' B) (Bound script) run configureThisAsMaster_() once; or\n' +
    ' C) Set Script Property MASTER_SPREADSHEET_ID to your sheet ID.'
  );
}

/** One-time: run inside the Master (bound script) to imprint its ID everywhere useful. */
function configureThisAsMaster_() {
  const active = SpreadsheetApp.getActive();
  if (!active) throw new Error('Run this from the Master spreadsheet (bound script).');
  const id = active.getId();

  PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', id);

  const sh = active.getSheetByName('Settings') || active.insertSheet('Settings');
  const v = sh.getDataRange().getValues();
  const hasHeader = v && v.length > 0 && String(v[0][0] || '').toLowerCase() === 'key';
  if (!hasHeader) sh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);

  // upsert MASTER_ID row
  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][0] || '').trim().toUpperCase() === 'MASTER_ID') { row = r + 1; break; }
  }
  if (row === -1) row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 2).setValues([['MASTER_ID', id]]);

  // refresh caches
  __MASTER_ID_CACHE = id;
  __MASTER_SS_CACHE = active;

  return { ok: true, master_id: id };
}

/** Utility: set the Script Property MASTER_SPREADSHEET_ID from code. */
function setMasterIdProperty_(id) {
  const clean = String(id || '').trim();
  if (!clean) throw new Error('Pass a valid spreadsheet ID.');
  PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', clean);
  __MASTER_ID_CACHE = clean;
  __MASTER_SS_CACHE = null; // force reopen on next call
  return { ok: true, master_id: clean };
}

/**
 * Ensure a sheet exists with the given headers.
 * If missing → create it. If existing → update headers in row 1 to match.
 */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  // Ensure headers (row 1) match exactly
  const want = (headers || []).map(String);
  const have = sh.getRange(1, 1, 1, Math.max(want.length, sh.getLastColumn())).getValues()[0].map(String);

  // Create/overwrite row 1
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, want.length).setValues([want]);
  } else {
    for (let c = 0; c < want.length; c++) {
      if ((have[c] || '').trim() !== want[c]) {
        sh.getRange(1, c + 1).setValue(want[c]);
      }
    }
  }

  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, Math.min(30, sh.getLastColumn())); } catch(e) {}

  return sh;
}

/** Convenience: always use the Master + guarantee headers. */
function getOrCreateMasterSheet_(name, headers) {
  return ensureSheet_(master_(), name, headers);
}
