
/******************************************************
 * 02_master_access.gs — Master workbook accessor
 * Depends on:
 * - MASTER_SPREADSHEET_ID, SHEETS (00_constants.gs)
 *
 * Exports:
 * - master_() : Spreadsheet object
 * - getOrCreateMasterSheet_(name,hdrs) : Sheet with headers (created if missing)
 ******************************************************/

/**
 * Return the Master spreadsheet object.
 * Uses MASTER_SPREADSHEET_ID from constants. Throws if unset.
 */
function master_() {
  if (!MASTER_SPREADSHEET_ID || MASTER_SPREADSHEET_ID === 'PASTE_MASTER_ID_HERE') {
    throw new Error('MASTER_SPREADSHEET_ID is not set in 00_constants.gs');
  }
  return SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
}

/**
 * Get a sheet by name, creating it with headers if missing.
 * If headers are supplied and sheet is new/empty, writes them in row 1.
 */
function getOrCreateMasterSheet_(name, headersOpt) {
  const ss = master_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (headersOpt && sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headersOpt.length).setValues([headersOpt]);
    sh.setFrozenRows(1);
  }
  return sh;
}
