/** ================================
 * 102_Exceptions.gs (QA Tracker)
 * Canonical “Exceptions” store
 * ================================
 *
 * Depends on:
 * - master_() (02_master_access.gs)
 * - Session.getActiveUser() (email)
 *
 * Tab: "Exceptions" (single canonical sheet)
 * Header (fixed order):
 * timestamp_utc | ref | analyst_email | analyst_id | date_iso |
 * start_ts | end_ts | minutes | category | reason | status |
 * leads_informed | applied_by | applied_ts_utc | applied_note |
 * team | org | extra
 */

var Exceptions = (function () {
  var SHEET_NAME = 'Exceptions';
  var HEADER = [
    'timestamp_utc','ref','analyst_email','analyst_id',
    'date_iso','start_ts','end_ts','minutes',
    'category','reason','status','leads_informed',
    'applied_by','applied_ts_utc','applied_note',
    'team','org','extra'
  ];

  /* ---------------- Sheet guardians ---------------- */
  function ss_(){ return master_(); }

  function sheet_(){
    var sh = ss_().getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss_().insertSheet(SHEET_NAME);
      sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
      sh.setFrozenRows(1);
      try { sh.autoResizeColumns(1, HEADER.length); } catch(e){}
      return sh;
    }
    ensureHeader_(sh);
    return sh;
  }

  function ensureHeader_(sh){
    var width = Math.max(HEADER.length, sh.getLastColumn() || 1);
    var current = sh.getRange(1,1,1,width).getValues()[0] || [];
    for (var i=0;i<HEADER.length;i++){
      if (String(current[i]||'') !== HEADER[i]) {
        sh.getRange(1,i+1).setValue(HEADER[i]);
      }
    }
    sh.setFrozenRows(1);
  }

  /* ---------------- Small utils ---------------- */
  function isoNowUtc_(){ return new Date().toISOString(); }
  function pad2_(n){ return (n<10?'0':'')+n; }
  function ymd_(d){ return d.getFullYear()+'-'+pad2_(d.getMonth()+1)+'-'+pad2_(d.getDate()); }

  // Accepts "YYYY-MM-DDTHH:mm[:ss]" (local) or "YYYY-MM-DD HH:mm[:ss]" strings; returns Date (local)
  function parseLocalIso_(isoLocal){
    if (!isoLocal) return null;
    var s = String(isoLocal).trim().replace('T',' ');
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return null;
    var y=+m[1], mo=+m[2]-1, d=+m[3], h=+m[4], mi=+m[5], se=+(m[6]||0);
    var dt = new Date(y,mo,d,h,mi,se);
    return isNaN(dt) ? null : dt;
  }

  function minutesBetween_(startIso, endIso){
    var s = parseLocalIso_(startIso), e = parseLocalIso_(endIso);
    if (!s || !e) return 0;
    return Math.round(Math.max(0, e.getTime()-s.getTime())/60000);
  }

  function genRef_(){
    var d = new Date();
    var refDay = ymd_(d).replace(/-/g,'');
    var rand = Math.floor(Math.random()*0x2710).toString(16).toUpperCase(); // 0..9999 → hex
    return 'EXC-' + refDay + '-' + ('0000'+rand).slice(-4);
  }

  function asMap_(row){
    var o={};
    for (var i=0;i<HEADER.length;i++) o[HEADER[i]] = (row[i]===undefined?'':row[i]);
    return o;
  }

  function findRowIndexByRef_(sh, ref){
    if (!ref) return -1;
    // Search the 'ref' column (B) only for speed.
    var last = sh.getLastRow();
    if (last < 2) return -1;
    var vals = sh.getRange(2,2,last-1,1).getValues(); // col B
    for (var i=0;i<vals.length;i++){
      if (String(vals[i][0]) === String(ref)) return (i+2);
    }
    return -1;
  }

  function email_(){
    try { return Session.getActiveUser().getEmail() || ''; } catch(e){ return ''; }
  }

  function string_(v){ return (v==null ? '' : String(v).trim()); }
  function clampMinutes_(n){ var m=Math.max(0,Math.round(Number(n)||0)); return m>0?m:0; }

  function sanitizeCategory_(c){
    var s=(c||'').toString().trim();
    if (!s) return 'Other';
    var allowed=['System Outage','Appointment','Training','Connectivity','Other'];
    if (allowed.indexOf(s)>=0) return s;
    // Normalise arbitrary text into Title case-ish
    return s.charAt(0).toUpperCase()+s.slice(1).toLowerCase();
  }

  function deriveAnalystIdFromEmail_(em){
    return String(em||'').split('@')[0] || '';
  }

  /* ---------------- Core API ---------------- */

  /**
   * Create exception from analyst submission.
   * payload: {
   * analyst_id?, date_iso?, start_ts, end_ts, category?, reason?, leads?,
   * team?, org?, __ua?
   * }
   */
  function createFromAnalyst(payload){
    payload = payload || {};
    var sh = sheet_();

    var me = email_();
    var ref = genRef_();

    var dateIso = string_(payload.date_iso);
    var startIso = string_(payload.start_ts);
    var endIso = string_(payload.end_ts);
    var mins = minutesBetween_(startIso, endIso);

    if (!dateIso) {
      var sd = parseLocalIso_(startIso) || new Date();
      dateIso = ymd_(sd);
    }
    if (mins <= 0) throw new Error('End time must be after start time.');

    var analystId = string_(payload.analyst_id) || deriveAnalystIdFromEmail_(me);

    var extra = {
      source: 'analyst',
      ua: String(payload.__ua || '')
    };

    var row = [
      isoNowUtc_(), // timestamp_utc
      ref, // ref
      string_(me), // analyst_email
      analystId, // analyst_id
      dateIso, // date_iso
      startIso, // start_ts (local ISO-like)
      endIso, // end_ts (local ISO-like)
      mins, // minutes
      sanitizeCategory_(payload.category), // category
      string_(payload.reason), // reason
      'NEW', // status
      Array.isArray(payload.leads) ? payload.leads.join(',') : string_(payload.leads), // leads_informed
      '', // applied_by
      '', // applied_ts_utc
      '', // applied_note
      string_(payload.team), // team
      string_(payload.org), // org
      JSON.stringify(extra) // extra
    ];

    sh.appendRow(row);
    return { ref: ref, minutes: mins, status: 'NEW' };
  }

  /**
   * Lead view with filters.
   * params: { sinceDays?, status?('ANY'|'NEW'|'ACK'|'APPLIED'|'REJECTED'|'CANCELED'), forDate?, analyst?, team? }
   */
  function listForLead(params){
    params = params || {};
    var sh = sheet_();
    var last = sh.getLastRow();
    if (last < 2) return { items: [] };

    var rows = sh.getRange(2,1,last-1,HEADER.length).getValues();
    var sinceDays = Number(params.sinceDays||7);
    var cutoff = new Date(); cutoff.setDate(cutoff.getDate()-sinceDays);
    var wantStatus = (params.status||'ANY').toUpperCase();

    var out=[];
    for (var i=0;i<rows.length;i++){
      var m=asMap_(rows[i]);

      if (wantStatus!=='ANY' && String(m.status||'').toUpperCase()!==wantStatus) continue;
      if (params.forDate && String(m.date_iso||'') !== String(params.forDate)) continue;
      if (params.analyst && String(m.analyst_email||'').toLowerCase() !== String(params.analyst).toLowerCase()) continue;
      if (params.team && String(m.team||'').toLowerCase() !== String(params.team).toLowerCase()) continue;

      var ts = new Date(m.timestamp_utc || rows[i][0] || new Date());
      if (ts < cutoff) continue;

      out.push(m);
    }
    out.sort(function(a,b){ return new Date(b.timestamp_utc) - new Date(a.timestamp_utc); });
    return { items: out };
  }

  function getByRef(ref){
    var sh = sheet_();
    var idx = findRowIndexByRef_(sh, ref);
    if (idx < 0) return null;
    var row = sh.getRange(idx,1,1,HEADER.length).getValues()[0];
    return asMap_(row);
  }

  /**
   * status ∈ {'NEW','ACK','APPLIED','REJECTED','CANCELED'}
   */
  function setStatus(ref,status,note){
    status = (status||'').toUpperCase();
    if (['NEW','ACK','APPLIED','REJECTED','CANCELED'].indexOf(status)<0){
      throw new Error('Invalid status: '+status);
    }
    var sh = sheet_();
    var i = findRowIndexByRef_(sh, ref);
    if (i < 0) throw new Error('Ref not found: '+ref);

    var row = sh.getRange(i,1,1,HEADER.length).getValues()[0];
    row[10] = status; // status
    if (note) row[14] = String(note); // applied_note
    if (status==='APPLIED' && !row[13]) row[13] = isoNowUtc_(); // applied_ts_utc
    if (status==='APPLIED' && !row[12]) row[12] = email_(); // applied_by

    sh.getRange(i,1,1,HEADER.length).setValues([row]);
    return { ref:ref, status:status };
  }

  /**
   * Mark as APPLIED; optional minutes adjustment.
   * opts: { minus_minutes?, applied_note?, cal_block_id? }
   */
  function markApplied(ref, opts){
    opts = opts || {};
    var sh = sheet_();
    var i = findRowIndexByRef_(sh, ref);
    if (i < 0) throw new Error('Ref not found: '+ref);

    var row = sh.getRange(i,1,1,HEADER.length).getValues()[0];
    var minutes = clampMinutes_(opts.minus_minutes != null ? opts.minus_minutes : row[7]);

    row[10] = 'APPLIED';
    row[12] = email_();
    row[13] = isoNowUtc_();
    row[14] = buildAppliedNote_(row, opts);
    if (opts.minus_minutes != null && opts.minus_minutes >= 0) row[7] = minutes;

    sh.getRange(i,1,1,HEADER.length).setValues([row]);
    return { ref: ref, status:'APPLIED', minutes: minutes };
  }

  function buildAppliedNote_(row, opts){
    var bits=[];
    if (opts.applied_note) bits.push(String(opts.applied_note));
    if (opts.cal_block_id) bits.push('calBlock='+opts.cal_block_id);
    return bits.join(' | ');
  }

  /** Simple text search across a few fields. */
  function search(q){
    q = String(q||'').toLowerCase();
    var sh = sheet_();
    var last = sh.getLastRow();
    if (last < 2) return { items: [] };

    var rows = sh.getRange(2,1,last-1,HEADER.length).getValues();
    var out=[];
    for (var i=0;i<rows.length;i++){
      var m = asMap_(rows[i]);
      var hay = [m.ref,m.analyst_email,m.analyst_id,m.category,m.reason,m.status].join(' ').toLowerCase();
      if (hay.indexOf(q)>=0) out.push(m);
    }
    out.sort(function(a,b){ return new Date(b.timestamp_utc) - new Date(a.timestamp_utc); });
    return { items: out };
  }

  return {
    createFromAnalyst: createFromAnalyst,
    listForLead: listForLead,
    getByRef: getByRef,
    setStatus: setStatus,
    markApplied: markApplied,
    search: search
  };
})();

/* -------- Optional wrappers (kept for compatibility) -------- */
function Exceptions_createFromAnalyst(payload){ return Exceptions.createFromAnalyst(payload); }
function Exceptions_listForLead(params){ return Exceptions.listForLead(params); }
function Exceptions_getByRef(ref){ return Exceptions.getByRef(ref); }
function Exceptions_setStatus(ref,status,note){ return Exceptions.setStatus(ref,status,note); }
function Exceptions_markApplied(ref,opts){ return Exceptions.markApplied(ref,opts); }
function Exceptions_search(q){ return Exceptions.search(q); }
