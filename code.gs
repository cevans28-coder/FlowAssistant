/******** Backup File

/***** 07_exceptions_helpers.gs
 * Extra helpers around Exceptions + Notifications for TL Console / Admin
 * Safe to add; no UI changes required.
 */

/* ------------------------------------------------------------
 * Internal resolvers (reuse patterns from file #6)
 * ------------------------------------------------------------ */
function _ex_resolve_() {
  if (typeof Exceptions === 'object' && Exceptions) {
    return {
      listForLead: (params) => Exceptions.listForLead(params),
      getByRef: (ref) => Exceptions.getByRef(ref),
      search: (q) => Exceptions.search(q),
      markApplied: (ref,o) => Exceptions.markApplied(ref,o),
      _alt: {
        listForLead: (p) => (typeof Exceptions_listForLead === 'function' ? Exceptions_listForLead(p) : null),
        getByRef: (r) => (typeof Exceptions_getByRef === 'function' ? Exceptions_getByRef(r) : null),
        search: (q) => (typeof Exceptions_search === 'function' ? Exceptions_search(q) : null),
        markApplied: (r,o)=> (typeof Exceptions_markApplied === 'function' ? Exceptions_markApplied(r,o) : null),
      }
    };
  }
  if (typeof QATracker === 'object' && QATracker) {
    // Prefix form
    if (typeof QATracker.Exceptions_listForLead === 'function' ||
        typeof QATracker.Exceptions_getByRef === 'function' ||
        typeof QATracker.Exceptions_search === 'function' ||
        typeof QATracker.Exceptions_markApplied === 'function') {
      return {
        listForLead: (params) => QATracker.Exceptions_listForLead(params),
        getByRef: (ref) => QATracker.Exceptions_getByRef(ref),
        search: (q) => QATracker.Exceptions_search(q),
        markApplied: (ref,o) => QATracker.Exceptions_markApplied(ref,o),
        _alt: {}
      };
    }
    // Namespaced object
    if (QATracker.Exceptions && typeof QATracker.Exceptions === 'object') {
      return {
        listForLead: (params) => QATracker.Exceptions.listForLead(params),
        getByRef: (ref) => QATracker.Exceptions.getByRef(ref),
        search: (q) => QATracker.Exceptions.search(q),
        markApplied: (ref,o) => QATracker.Exceptions.markApplied(ref,o),
        _alt: {}
      };
    }
  }
  throw new Error('Exceptions module not found (local or QATracker).');
}

function _ntf_resolve_() {
  if (typeof Notifications === 'object' && Notifications) {
    return {
      list: (p) => Notifications.list(p),
      _alt: { list: (p) => (typeof Notifications_list === 'function' ? Notifications_list(p) : null) }
    };
  }
  if (typeof QATracker === 'object' && QATracker) {
    if (typeof QATracker.Notifications_list === 'function') {
      return { list: (p)=> QATracker.Notifications_list(p), _alt:{} };
    }
    if (QATracker.Notifications && typeof QATracker.Notifications.list === 'function') {
      return { list: (p)=> QATracker.Notifications.list(p), _alt:{} };
    }
  }
  throw new Error('Notifications module not found (local or QATracker).');
}

/* ------------------------------------------------------------
 * Public: List exceptions (TL-friendly)
 * params: {
 * sinceDays?: number = 30,
 * status?: 'NEW'|'ACK'|'APPLIED'|'REJECTED'|'CANCELED'|'ANY' = 'ANY',
 * forDate?: 'YYYY-MM-DD',
 * analyst?: string (email or id),
 * team?: string
 * }
 * Returns { items: [...] } as provided by Exceptions.listForLead
 * ------------------------------------------------------------ */
function tlListExceptions(params) {
  if (!isTeamLead_()) throw new Error('Not authorised');
  params = params || {};
  if (params.status) params.status = String(params.status).toUpperCase();
  if (!('sinceDays' in params)) params.sinceDays = (params.status === 'ANY' ? 180 : 30);

  var ex = _ex_resolve_();
  try {
    return ex.listForLead(params) || { items: [] };
  } catch (e) {
    if (ex._alt.listForLead) return ex._alt.listForLead(params) || { items: [] };
    throw e;
  }
}

/* ------------------------------------------------------------
 * Public: Get one exception by ref
 * ------------------------------------------------------------ */
function tlGetExceptionByRef(ref) {
  if (!isTeamLead_()) throw new Error('Not authorised');
  if (!ref) throw new Error('Missing ref');
  var ex = _ex_resolve_();
  try {
    return ex.getByRef(ref);
  } catch (e) {
    if (ex._alt.getByRef) return ex._alt.getByRef(ref);
    throw e;
  }
}

/* ------------------------------------------------------------
 * Public: Free-text search exceptions
 * ------------------------------------------------------------ */
function tlSearchExceptions(q) {
  if (!isTeamLead_()) throw new Error('Not authorised');
  var ex = _ex_resolve_();
  try {
    return ex.search(q || '');
  } catch (e) {
    if (ex._alt.search) return ex._alt.search(q || '');
    throw e;
  }
}

/* ------------------------------------------------------------
 * Public: Convenience — list TL notification history
 * days: default 180
 * Returns { items: [...] } from Notifications.list
 * ------------------------------------------------------------ */
function tlListNotificationsHistory(days) {
  if (!isTeamLead_()) throw new Error('Not authorised');
  var svc = _ntf_resolve_();
  var p = { audience:'TL', status:'ANY', sinceDays: Number(days || 180) };
  try {
    var raw = svc.list(p);
    return Array.isArray(raw) ? { items: raw } : (raw || { items: [] });
  } catch (e) {
    if (svc._alt.list) {
      var alt = svc._alt.list(p);
      return Array.isArray(alt) ? { items: alt } : (alt || { items: [] });
    }
    throw e;
  }
}

/* ------------------------------------------------------------
 * Public: Apply an exception then rebuild metrics for that analyst/day
 * opts: { applied_note?:string, cal_block_id?:string, minus_minutes?:number }
 * Returns: { applied:{ref,status,minutes}, rebuilt:any }
 * ------------------------------------------------------------ */
function tlApplyExceptionAndRebuild(analystId, dateISO, ref, opts) {
  if (!isTeamLead_()) throw new Error('Not authorised');
  if (!ref) throw new Error('Missing exception REF');
  var ex = _ex_resolve_();

  // 1) Mark APPLIED
  var applied;
  try {
    applied = ex.markApplied(ref, opts || {});
  } catch (e) {
    if (ex._alt.markApplied) applied = ex._alt.markApplied(ref, opts || {});
    else throw e;
  }

  // 2) Rebuild metrics (best-effort)
  var rebuilt = null;
  try {
    if (typeof QATracker === 'object' && QATracker && typeof QATracker.buildMetricsForAnalystDate === 'function') {
      rebuilt = QATracker.buildMetricsForAnalystDate(analystId || '', dateISO || '');
    } else if (typeof tlBuildMetricsForAnalystProxy === 'function') {
      rebuilt = tlBuildMetricsForAnalystProxy(analystId || '', dateISO || '');
    }
  } catch (e2) {
    // non-fatal; still return `applied`
    rebuilt = { ok:false, error:String(e2 && e2.message || e2) };
  }

  return { applied: applied, rebuilt: rebuilt };
}

/* ------------------------------------------------------------
 * Public: Health check for admin/debug
 * - Confirms access, sheet presence, and module resolution
 * ------------------------------------------------------------ */
function tlHealthCheck() {
  var ok = true; var notes = [];

  // permission
  try {
    if (!isTeamLead_()) throw new Error('Not authorised');
    notes.push('Auth: OK');
  } catch (e) {
    ok = false; notes.push('Auth: FAIL — ' + (e && e.message));
  }

  // master + key sheets
  try {
    var ss = master_();
    var needed = ['DailyMetrics','DailyTypeSummary','Live','StatusLogs','Analysts'];
    var missing = [];
    needed.forEach(function(n){
      if (!ss.getSheetByName(n)) missing.push(n);
    });
    if (missing.length) { notes.push('Sheets missing: ' + missing.join(', ')); }
    else notes.push('Sheets: OK');
  } catch (e) {
    ok = false; notes.push('Master/Sheets: FAIL — ' + (e && e.message));
  }

  // modules
  try {
    _ex_resolve_(); notes.push('Exceptions: OK');
  } catch (e) {
    ok = false; notes.push('Exceptions: FAIL — ' + (e && e.message));
  }
  try {
    _ntf_resolve_(); notes.push('Notifications: OK');
  } catch (e) {
    ok = false; notes.push('Notifications: FAIL — ' + (e && e.message));
  }

  return { ok: ok, notes: notes };
}
