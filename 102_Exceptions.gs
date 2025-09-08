/** ================================
 * 102_Exceptions.gs (QA Tracker)
 * Canonical “Exceptions” store
 * ================================
 */

function getMasterSS_() {
  var id = PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('MASTER_SHEET_ID not set and no active spreadsheet. Run setMasterSheetId(id).');
  return ss;
}
function setMasterSheetId(id) {
  PropertiesService.getScriptProperties().setProperty('MASTER_SHEET_ID', id);
  return { ok:true, id:id };
}

var Exceptions = (function () {
  var SHEET_NAME = 'Exceptions';
  var HEADER = [
    'timestamp_utc','ref','analyst_email','analyst_id',
    'date_iso','start_ts','end_ts','minutes',
    'category','reason','status','leads_informed',
    'applied_by','applied_ts_utc','applied_note',
    'team','org','extra'
  ];

  function ss_(){ return getMasterSS_(); }
  function sheet_(){
    var sh = ss_().getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss_().insertSheet(SHEET_NAME);
      sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
    } else {
      ensureHeader_(sh);
    }
    return sh;
  }
  function ensureHeader_(sh){
    var current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), HEADER.length)).getValues()[0];
    for (var i=0;i<HEADER.length;i++){
      if (current[i] !== HEADER[i]) {
        sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
        break;
      }
    }
  }
  function isoNowUtc_(){ return new Date().toISOString(); }
  function pad2_(n){ return (n<10?'0':'')+n; }
  function ymd_(d){ return d.getFullYear()+'-'+pad2_(d.getMonth()+1)+'-'+pad2_(d.getDate()); }
  function parseLocalIso_(isoLocal){
    if (!isoLocal) return null;
    var parts = isoLocal.split(/[T ]/);
    var d = parts[0].split('-'); var t = (parts[1]||'00:00:00').split(':');
    return new Date(Number(d[0]),Number(d[1])-1,Number(d[2]),Number(t[0]||0),Number(t[1]||0),Number(t[2]||0));
  }
  function minutesBetween_(startIso, endIso){
    var s = parseLocalIso_(startIso), e = parseLocalIso_(endIso);
    if (!s || !e) return 0;
    return Math.round(Math.max(0, e.getTime()-s.getTime())/60000);
  }
  function genRef_(){
    var d = new Date();
    var refDay = ymd_(d).replace(/-/g,'');
    var rand = Math.floor(Math.random()*0x2710).toString(16).toUpperCase();
    return 'EXC-' + refDay + '-' + ('0000'+rand).slice(-4);
  }
  function asMap_(row){ var o={}; for (var i=0;i<HEADER.length;i++) o[HEADER[i]]=(row[i]===undefined?'':row[i]); return o; }
  function findRowIndexByRef_(sh, ref){
    var vals = sh.getRange(2,2,Math.max(0,sh.getLastRow()-1),1).getValues();
    for (var i=0;i<vals.length;i++){ if (String(vals[i][0])===String(ref)) return (i+2); }
    return -1;
  }
  function email_(){ try { return Session.getActiveUser().getEmail() || ''; } catch(e){ return ''; } }
  function string_(v){ return (v==null?'':String(v).trim()); }
  function clampMinutes_(n){ var m=Math.max(0,Math.round(Number(n)||0)); return m>0?m:0; }
  function sanitizeCategory_(c){
    var s=(c||'').toString().trim(); if (!s) return 'Other';
    var allowed=['System Outage','Appointment','Training','Connectivity','Other'];
    if (allowed.indexOf(s)>=0) return s;
    return s.charAt(0).toUpperCase()+s.slice(1).toLowerCase();
  }
  function deriveAnalystIdFromEmail_(email){
    return String(email||'').split('@')[0] || '';
  }

  // -------- core --------

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
      isoNowUtc_(),
      ref,
      string_(me),
      analystId,
      dateIso,
      startIso,
      endIso,
      mins,
      sanitizeCategory_(payload.category),
      string_(payload.reason),
      'NEW',
      Array.isArray(payload.leads) ? payload.leads.join(',') : string_(payload.leads),
      '', '', '',
      string_(payload.team),
      string_(payload.org),
      JSON.stringify(extra)
    ];

    sh.appendRow(row);
    return { ref: ref, minutes: mins, status: 'NEW' };
  }

  function listForLead(params){
    params = params || {};
    var sh = sheet_();
    var rows = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), HEADER.length).getValues();
    var sinceDays = Number(params.sinceDays||7);
    var cutoff = new Date(); cutoff.setDate(cutoff.getDate()-sinceDays);
    var wantStatus = (params.status||'ANY').toUpperCase();

    var out=[];
    for (var i=0;i<rows.length;i++){
      var m=asMap_(rows[i]);
      if (wantStatus!=='ANY' && m.status!==wantStatus) continue;
      if (params.forDate && m.date_iso!==params.forDate) continue;
      if (params.analyst && String(m.analyst_email).toLowerCase()!==String(params.analyst).toLowerCase()) continue;
      if (params.team && String(m.team||'').toLowerCase()!==String(params.team).toLowerCase()) continue;

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

  function setStatus(ref,status,note){
    status = (status||'').toUpperCase();
    if (['NEW','ACK','APPLIED','REJECTED','CANCELED'].indexOf(status)<0){
      throw new Error('Invalid status: '+status);
    }
    var sh = sheet_();
    var i = findRowIndexByRef_(sh, ref);
    if (i < 0) throw new Error('Ref not found: '+ref);

    var row = sh.getRange(i,1,1,HEADER.length).getValues()[0];
    row[10] = status;
    if (note) row[14] = String(note);
    if (status==='APPLIED' && !row[13]) row[13] = isoNowUtc_();
    if (status==='APPLIED' && !row[12]) row[12] = email_();

    sh.getRange(i,1,1,HEADER.length).setValues([row]);
    return { ref:ref, status:status };
  }

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

  function search(q){
    q = String(q||'').toLowerCase();
    var sh = sheet_();
    var rows = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), HEADER.length).getValues();
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

/* Optional wrappers */
function Exceptions_createFromAnalyst(payload){ return Exceptions.createFromAnalyst(payload); }
function Exceptions_listForLead(params){ return Exceptions.listForLead(params); }
function Exceptions_getByRef(ref){ return Exceptions.getByRef(ref); }
function Exceptions_setStatus(ref,status,note){ return Exceptions.setStatus(ref,status,note); }
function Exceptions_markApplied(ref,opts){ return Exceptions.markApplied(ref,opts); }
function Exceptions_search(q){ return Exceptions.search(q); }
