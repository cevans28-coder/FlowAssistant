/******************************************************
 * Smart Master resolver — works in both contexts
 * Order of precedence:
 * 1) Global constant MASTER_SPREADSHEET_ID (if set & not placeholder)
 * 2) Active spreadsheet (when running as a bound script)
 * 3) Script Properties: MASTER_SPREADSHEET_ID (configured once)
 * 4) Settings tab key MASTER_ID (optional fallback)
 ******************************************************/

/** Try to read a usable Master ID from multiple places. */
function getMasterId_() {
  // 1) Global constant (from 00_constants.gs)
  try {
    if (typeof MASTER_SPREADSHEET_ID !== 'undefined') {
      const v = String(MASTER_SPREADSHEET_ID || '').trim();
      if (v && v !== 'PASTE_YOUR_MASTER_SHEET_ID_HERE') return v;
    }
  } catch (e) { /* ignore */ }

  // 2) If we are a bound project, the active spreadsheet *is* the master
  try {
    const as = SpreadsheetApp.getActive();
    if (as) return as.getId();
  } catch (e) { /* ignore */ }

  // 3) Script Properties (works for both bound and standalone projects)
  try {
    const prop = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');
    if (prop && prop.trim()) return prop.trim();
  } catch (e) { /* ignore */ }

  // 4) Settings tab fallback (if running bound, or if an ID is already known)
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
            if (id) return id;
          }
        }
      }
    }
  } catch (e) { /* ignore */ }

  return ''; // not found
}

/** Open the Master spreadsheet with robust fallbacks + clear error if missing. */
function master_() {
  const id = getMasterId_();
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {
      throw new Error('MASTER_SPREADSHEET_ID appears set but cannot be opened. Check sharing/ID.');
    }
  }

  // Nothing resolved: give concrete next steps
  throw new Error(
    'MASTER_SPREADSHEET_ID is not configured.\n\n' +
    'Fix it by doing one of the following:\n' +
    ' A) In 00_constants.gs, set const MASTER_SPREADSHEET_ID = "YOUR_SHEET_ID";\n' +
    ' B) (Bound script only) Run configureThisAsMaster_() once; or\n' +
    ' C) Set Script Property MASTER_SPREADSHEET_ID to your sheet ID.'
  );
}

/** One-time: run inside the Master (bound script) to imprint its ID everywhere useful. */
function configureThisAsMaster_() {
  const active = SpreadsheetApp.getActive();
  if (!active) throw new Error('Run this from the Master spreadsheet (bound script).');
  const id = active.getId();

  // Save as Script Property for this project
  PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', id);

  // Also store in Settings tab for human reference and optional fallback
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

  return { ok: true, master_id: id };
}

/** Utility: set the Script Property MASTER_SPREADSHEET_ID from code. */
function setMasterIdProperty_(id) {
  const clean = String(id || '').trim();
  if (!clean) throw new Error('Pass a valid spreadsheet ID.');
  PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', clean);
  return { ok: true, master_id: clean };
}
