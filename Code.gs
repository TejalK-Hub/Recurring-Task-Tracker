/**
 * Recurring Task Tracking & Escalation Prototype
 * Google Sheets + Google Apps Script (container-bound script)
 *
 * HOW TO USE
 *  1. Create a blank Google Sheet > Extensions > Apps Script > paste this file as Code.gs > Save.
 *  2. Project Settings: set time zone to Asia/Kolkata (optional but recommended).
 *  3. Run setupSystem() once from the editor and approve permissions (see Documentation.md).
 *  4. Use the "Task System" menu in the sheet.
 *
 * DESIGN RULES
 *  - Nothing here ever deletes or clears Cycles, Completions or Audit_Log rows.
 *  - Demo/test records use Task_IDs starting with DEMO_ and are ignored by the real engine,
 *    the Dashboard and real alerts. Demo emails always go to TEST_EMAIL.
 *  - Identity: GOOGLE_EMAIL only when the editor's email is available AND matches the task's
 *    Employee_Email; otherwise ASSIGNEE_FALLBACK (NOT verified identity).
 */

/**
 * @OnlyCurrentDoc
 */

// ---------------------------------------------------------------- constants
var TZ = 'Asia/Kolkata';
var TZ_OFFSET = '+05:30'; // IST has no daylight saving
var DEMO_PREFIX = 'DEMO_';
var MAXROWS = 1000;
var SN = {TM: 'Task_Master', CY: 'Cycles', DB: 'Dashboard', CF: 'Config', CP: 'Completions', AL: 'Audit_Log', TC: 'Test_Cases'};

var TM_H = ['Task_ID', 'Task_Name', 'Assigned_To', 'Employee_Email', 'Frequency_Days', 'Schedule_Mode', 'Proof_Type', 'Due_Time', 'Start_Date', 'Active', 'Notes'];
var CY_H = ['Cycle_ID', 'Task_ID', 'Task_Name', 'Assigned_To', 'Due_Date', 'Due_Time', 'Status', 'Proof_Type', 'Required_Proof',
  'Ticket_ID_or_Proof', 'Done', 'Completed_By', 'Completed_At', 'Identity_Method', 'Timeliness',
  'Manager_Alert_Sent_At', 'Manager_Acknowledged', 'Acknowledged_By', 'Acknowledged_At', 'Manager_Note',
  'Leadership_Escalated_At', 'Escalation_Reason', 'Alert_Status',
  'Cycle_No', 'Schedule_Date', 'Due_At', 'Created_At', 'Previous_Cycle_ID'];
var CP_H = ['Completion_ID', 'Cycle_ID', 'Task_ID', 'Task_Name', 'Cycle_No', 'Assigned_To', 'Completed_By', 'Completed_At', 'Due_At', 'Timeliness', 'Ticket_ID_or_Proof', 'Identity_Method'];
var AL_H = ['Log_ID', 'Timestamp', 'Event_Type', 'Cycle_ID', 'Task_ID', 'Actor', 'Details', 'Run_ID'];
var TC_H = ['Test_ID', 'Scenario', 'Setup/Action', 'Expected_Result', 'Actual_Result', 'Status', 'Evidence/Notes'];

function idx_(h) { var o = {}; h.forEach(function (n, i) { o[n] = i + 1; }); return o; }
var TM = idx_(TM_H), CY = idx_(CY_H), CP = idx_(CP_H);

var CFG_DEFAULTS = [
  ['TIMEZONE', 'Asia/Kolkata', 'Informational. Script is fixed to IST (+05:30).'],
  ['WORKING_DAYS', 'Mon,Tue,Wed,Thu,Fri,Sat', 'Sunday is non-working.'],
  ['HOLIDAYS', '', 'Optional. Comma-separated dates, format yyyy-MM-dd'],
  ['DEFAULT_DUE_TIME', '17:00', 'Used when a task has no Due_Time (24h HH:mm)'],
  ['MANAGER_NAME', 'Manager', 'Used as fallback name for acknowledgements'],
  ['MANAGER_EMAIL', 'manager@example.com', 'PLACEHOLDER - replace with real manager email'],
  ['LEADERSHIP_EMAIL', 'leadership@example.com', 'PLACEHOLDER - replace with real leadership email'],
  ['TEST_EMAIL', '', 'All TEST-mode and demo emails go here. Configure manually before testing.'],
  ['EMAIL_MODE', 'OFF', 'TEST = redirect all emails to TEST_EMAIL | LIVE = real recipients | OFF = send nothing'],
  ['MANAGER_ALERT_DELAY_MIN', '30', 'Minutes after the task deadline before the manager is alerted (5:00 PM + 30 = 5:30 PM)'],
  ['ACK_DEADLINE_TIME', '10:30', 'Manager must acknowledge by this time on the next working day'],
  ['RESOLUTION_DEADLINE_TIME', '17:00', 'Task must be resolved by this time on the next working day'],
  ['ENGINE_INTERVAL_MIN', '15', 'Allowed: 1, 5, 10, 15, 30. Re-run Setup after changing.'],
  ['DEMO_NOW_OVERRIDE', '', 'Leave BLANK normally. Format yyyy-MM-dd HH:mm. Overrides the clock for manual SLA demos.']
];

// ---------------------------------------------------------------- small helpers
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(n) { return ss_().getSheetByName(n); }
function toast_(m) { try { ss_().toast(m, 'Task System', 8); } catch (e) { } }
function pad_(n, w) { var s = String(n); while (s.length < w) s = '0' + s; return s; }
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f); }
function ds_(v) { return (v instanceof Date) ? fmt_(v, 'yyyy-MM-dd') : String(v == null ? '' : v).trim(); }
function ts_(v, def) {
  if (v instanceof Date) return fmt_(v, 'HH:mm');
  var s = String(v == null ? '' : v).trim(); if (!s) return def;
  var m = s.match(/^(\d{1,2}):(\d{2})/); return m ? pad_(m[1], 2) + ':' + m[2] : def;
}
function parseAt_(dateStr, timeStr) { return new Date(dateStr + 'T' + timeStr + ':00' + TZ_OFFSET); }
function addDays_(s, n) { var p = s.split('-'); return Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)), 'UTC', 'yyyy-MM-dd'); }
var DAYN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
function dow_(s) { var p = s.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay(); }
function isWorking_(s, cfg) { return cfg.workDays.indexOf(DAYN[dow_(s)]) > -1 && cfg.holidays.indexOf(s) < 0; }
function nextWorking_(s, cfg) { var d = addDays_(s, 1), g = 0; while (!isWorking_(d, cfg) && g++ < 400) d = addDays_(d, 1); return d; }
function onOrAfterWorking_(s, cfg) { var g = 0; while (!isWorking_(s, cfg) && g++ < 400) s = addDays_(s, 1); return s; }
function prevWorking_(s, cfg) { var g = 0; while (!isWorking_(s, cfg) && g++ < 400) s = addDays_(s, -1); return s; }
function isDemoId_(id) { return String(id).indexOf(DEMO_PREFIX) === 0; }
function dailyId_(tid, sched) { return tid + '-' + sched.replace(/-/g, ''); }

function computeStatus_(dueDate, dueAt, now) {
  if (now.getTime() > dueAt.getTime()) return 'Overdue';
  if (fmt_(now, 'yyyy-MM-dd') >= dueDate) return 'Due';
  return 'Scheduled';
}
function proofText_(p) {
  return String(p).toUpperCase() === 'TICKET_ID'
    ? 'Enter the portal Ticket ID in Ticket_ID_or_Proof, then tick Done'
    : 'Share the update in the WhatsApp group, optionally note a message reference, then tick Done';
}

function appendRows_(sh, rows) {
  if (!rows.length) return 0;
  var r = sh.getLastRow() + 1, need = r + rows.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(r, 1, rows.length, rows[0].length).setValues(rows);
  return r;
}

// ---------------------------------------------------------------- config
function getCfg_() {
  var sh = sh_(SN.CF), o = {};
  if (sh && sh.getLastRow() > 0) {
    sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) { if (r[0]) o[String(r[0]).trim()] = r[1]; });
  }
  function str(k, d) { var v = o[k]; return (v === undefined || v === '' || v === null) ? d : String(v).trim(); }
  var c = {};
  c.workDays = str('WORKING_DAYS', 'Mon,Tue,Wed,Thu,Fri,Sat').split(',').map(function (x) {
    x = x.trim().substr(0, 3); return x.charAt(0).toUpperCase() + x.substr(1).toLowerCase();
  }).filter(function (x) { return DAYN.indexOf(x) > -1; });
  if (!c.workDays.length) c.workDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  c.holidays = (o.HOLIDAYS instanceof Date ? [ds_(o.HOLIDAYS)] : str('HOLIDAYS', '').split(',')).map(function (x) { return x.trim(); }).filter(String);
  c.defaultDue = ts_(o.DEFAULT_DUE_TIME, '17:00');
  c.managerName = str('MANAGER_NAME', 'Manager');
  c.managerEmail = str('MANAGER_EMAIL', '');
  c.leadEmail = str('LEADERSHIP_EMAIL', '');
  c.testEmail = str('TEST_EMAIL', '');
  c.mode = str('EMAIL_MODE', 'OFF').toUpperCase();
  c.delay = Number(str('MANAGER_ALERT_DELAY_MIN', '30')); if (isNaN(c.delay)) c.delay = 30;
  c.ackTime = ts_(o.ACK_DEADLINE_TIME, '10:30');
  c.resTime = ts_(o.RESOLUTION_DEADLINE_TIME, '17:00');
  c.interval = Number(str('ENGINE_INTERVAL_MIN', '15')) || 15;
  c.override = o.DEMO_NOW_OVERRIDE;
  return c;
}
function getNow_(cfg) {
  var v = cfg.override;
  if (v instanceof Date) return v;
  var m = String(v || '').trim().match(/^(\d{4}-\d{2}-\d{2})[ T](\d{1,2}:\d{2})$/);
  if (m) return parseAt_(m[1], pad_(m[2].split(':')[0], 2) + ':' + m[2].split(':')[1]);
  return new Date();
}

// ---------------------------------------------------------------- audit log (buffered)
var LOGBUF_ = [];
function log_(type, cycleId, taskId, actor, details, runId) {
  LOGBUF_.push(['L' + Utilities.getUuid().replace(/-/g, '').substr(0, 10), new Date(), type, cycleId || '', taskId || '', actor || 'SYSTEM', String(details || '').substr(0, 1500), runId || '']);
}
function flushLog_() {
  if (!LOGBUF_.length) return;
  var rows = LOGBUF_; LOGBUF_ = [];
  appendRows_(sh_(SN.AL), rows);
}

// ---------------------------------------------------------------- data access
function getTasks_() {
  var sh = sh_(SN.TM), n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, TM_H.length).getValues().map(function (r, i) {
    return {
      row: i + 2, id: String(r[0]).trim(), name: r[1], who: r[2], email: String(r[3] || '').trim(),
      freq: Number(r[4]) || 1, mode: String(r[5]).trim().toUpperCase(), proof: String(r[6]).trim().toUpperCase(),
      due: ts_(r[7], ''), start: ds_(r[8]), active: (r[9] === true || String(r[9]).toUpperCase() === 'TRUE')
    };
  }).filter(function (t) { return t.id; });
}
function taskMap_() { var m = {}; getTasks_().forEach(function (t) { m[t.id] = t; }); return m; }
function getTaskById_(id) { return taskMap_()[id] || null; }
function allCycles_() {
  var sh = sh_(SN.CY), n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, CY_H.length).getValues().filter(function (r) { return r[0]; });
}
function loadIndex_() {
  var ix = {ids: {}, cnt: {}, open: {}, last: {}};
  allCycles_().forEach(function (r) {
    var id = r[CY.Cycle_ID - 1], tid = r[CY.Task_ID - 1];
    ix.ids[id] = true; ix.cnt[tid] = (ix.cnt[tid] || 0) + 1;
    if (r[CY.Status - 1] !== 'Completed') ix.open[tid] = (ix.open[tid] || 0) + 1;
    var no = Number(r[CY.Cycle_No - 1]) || 0;
    if (!ix.last[tid] || no > ix.last[tid].no) ix.last[tid] = {no: no, id: id, status: r[CY.Status - 1], completedAt: r[CY.Completed_At - 1]};
  });
  return ix;
}

// ---------------------------------------------------------------- cycle creation
function cycleRow_(t, cfg, cycleId, dueDate, cycleNo, sched, prevId, now) {
  var dt = t.due || cfg.defaultDue, dueAt = parseAt_(dueDate, dt), r = [];
  for (var i = 0; i < CY_H.length; i++) r.push('');
  function set(k, v) { r[CY[k] - 1] = v; }
  set('Cycle_ID', cycleId); set('Task_ID', t.id); set('Task_Name', t.name); set('Assigned_To', t.who);
  set('Due_Date', dueDate); set('Due_Time', dt); set('Status', computeStatus_(dueDate, dueAt, now));
  set('Proof_Type', t.proof); set('Required_Proof', proofText_(t.proof));
  set('Done', false); set('Manager_Acknowledged', false); set('Alert_Status', 'None');
  set('Cycle_No', cycleNo); set('Schedule_Date', sched); set('Due_At', dueAt); set('Created_At', now);
  set('Previous_Cycle_ID', prevId || '');
  return r;
}
function createDailyIfMissing_(t, cfg, sched, now, ix, runId) {
  var id = dailyId_(t.id, sched);
  if (ix.ids[id]) return null; // unique key Task_ID + Schedule_Date
  var no = (ix.cnt[t.id] || 0) + 1;
  appendRows_(sh_(SN.CY), [cycleRow_(t, cfg, id, sched, no, sched, '', now)]);
  ix.ids[id] = true; ix.cnt[t.id] = no; ix.open[t.id] = (ix.open[t.id] || 0) + 1;
  log_('CYCLE_CREATED', id, t.id, 'SYSTEM', 'Daily cycle for ' + sched, runId);
  return id;
}
function createRollingNext_(t, cfg, prevId, prevNo, completedAt, now, ix, runId) {
  var id = t.id + '-C' + pad_(prevNo + 1, 3);
  if (ix.ids[id] || (ix.open[t.id] || 0) > 0) return null; // duplicate or an open cycle already exists
  var base = (completedAt instanceof Date) ? completedAt : now;
  var due = onOrAfterWorking_(addDays_(fmt_(base, 'yyyy-MM-dd'), t.freq), cfg);
  appendRows_(sh_(SN.CY), [cycleRow_(t, cfg, id, due, prevNo + 1, due, prevId, now)]);
  ix.ids[id] = true; ix.cnt[t.id] = (ix.cnt[t.id] || 0) + 1; ix.open[t.id] = 1;
  log_('CYCLE_CREATED', id, t.id, 'SYSTEM', 'Rolling cycle due ' + due + ' (completion date + ' + t.freq + ' days, shifted to working day)', runId);
  return id;
}
function generateCycles_(cfg, now, st, runId) {
  var ix = loadIndex_(), today = fmt_(now, 'yyyy-MM-dd');
  getTasks_().forEach(function (t) {
    if (!t.active || isDemoId_(t.id)) return;
    try {
      if (t.mode === 'CALENDAR_DAILY') {
        if (!isWorking_(today, cfg) || (t.start && today < t.start)) return;
        if (createDailyIfMissing_(t, cfg, today, now, ix, runId)) st.created++; else st.dup++;
      } else if (t.mode === 'ROLLING') {
        var last = ix.last[t.id];
        if (!last) {
          if (!t.start) { st.errors++; log_('AUTOMATION_ERROR', '', t.id, 'SYSTEM', 'Rolling task has no Start_Date', runId); return; }
          var due = onOrAfterWorking_(t.start, cfg), id = t.id + '-C001';
          if (ix.ids[id]) { st.dup++; return; }
          appendRows_(sh_(SN.CY), [cycleRow_(t, cfg, id, due, 1, due, '', now)]);
          ix.ids[id] = true; ix.cnt[t.id] = 1; ix.open[t.id] = 1; st.created++;
          log_('CYCLE_CREATED', id, t.id, 'SYSTEM', 'Initial rolling cycle due ' + due, runId);
        } else if (!(ix.open[t.id] > 0) && last.status === 'Completed') {
          // repair: completed cycle without a successor (e.g. interrupted run)
          if (createRollingNext_(t, cfg, last.id, last.no, last.completedAt, now, ix, runId)) st.created++;
        } else st.dup++;
      } else { st.errors++; log_('AUTOMATION_ERROR', '', t.id, 'SYSTEM', 'Unknown Schedule_Mode: ' + t.mode, runId); }
    } catch (err) { st.errors++; log_('AUTOMATION_ERROR', '', t.id, 'SYSTEM', 'generate: ' + err, runId); }
  });
}

// ---------------------------------------------------------------- employee completion (Done checkbox)
function activeEmail_(e) {
  var m = '';
  try { m = Session.getActiveUser().getEmail() || ''; } catch (x) { }
  if (!m && e && e.user) { try { m = e.user.getEmail() || ''; } catch (x) { } }
  return m;
}

function handleDone_(sh, row, now, ctx, ticked) {
  var cfg = ctx.cfg, r = sh.getRange(row, 1, 1, CY_H.length).getValues()[0];
  function g(k) { return r[CY[k] - 1]; }
  var cid = g('Cycle_ID'), tid = g('Task_ID'), assignee = g('Assigned_To');
  if (!cid) return null;
  var doneCell = sh.getRange(row, CY.Done);
  if (g('Status') === 'Completed') {
    if (!ticked) {
      doneCell.setValue(true);
      log_('COMPLETION_REJECTED', cid, tid, assignee, 'Cycle already completed; completed cycles are locked.');
      return {ok: false, msg: 'This cycle is already completed and locked.'};
    }
    return {ok: true, dup: true};
  }
  if (!ticked) return null;
  function reject(why) {
    doneCell.setValue(false);
    log_('COMPLETION_REJECTED', cid, tid, assignee, why);
    return {ok: false, msg: 'Not recorded: ' + why};
  }
  var dueDate = ds_(g('Due_Date')), dueAt = parseAt_(dueDate, ts_(g('Due_Time'), cfg.defaultDue));
  if (computeStatus_(dueDate, dueAt, now) === 'Scheduled') return reject('Cycle is not due until ' + dueDate + '. Early completion is not allowed.');
  var proof = String(g('Ticket_ID_or_Proof') == null ? '' : g('Ticket_ID_or_Proof')).trim();
  if (String(g('Proof_Type')).toUpperCase() === 'TICKET_ID' && !/^[A-Za-z0-9][A-Za-z0-9_\-#\/]{2,39}$/.test(proof)) {
    return reject('Enter a valid Ticket ID (3-40 letters/numbers/-/_) in Ticket_ID_or_Proof before ticking Done.');
  }
  // identity
  var task = getTaskById_(tid), email = String(ctx.email || '').toLowerCase();
  var method = 'ASSIGNEE_FALLBACK', by = assignee;
  if (email && task && task.email && email === task.email.toLowerCase()) { method = 'GOOGLE_EMAIL'; by = assignee + ' <' + email + '>'; }
  var timely = now.getTime() <= dueAt.getTime() ? 'On Time' : 'Late';
  var hadAlert = !!g('Manager_Alert_Sent_At') || !!g('Leadership_Escalated_At');
  sh.getRange(row, CY.Status).setValue('Completed');
  sh.getRange(row, CY.Completed_By, 1, 4).setValues([[by, now, method, timely]]);
  sh.getRange(row, CY.Alert_Status).setValue(hadAlert ? 'Resolved' : 'None');
  SpreadsheetApp.flush();
  // ledger: one entry per Cycle_ID
  var cp = sh_(SN.CP), n = cp.getLastRow() - 1, exists = false;
  if (n > 0) exists = cp.getRange(2, CP.Cycle_ID, n, 1).getValues().some(function (x) { return x[0] === cid; });
  if (exists) log_('DUPLICATE_PREVENTED', cid, tid, 'SYSTEM', 'Completion ledger entry already exists; not added again.');
  else appendRows_(cp, [['CMP-' + cid, cid, tid, g('Task_Name'), g('Cycle_No'), assignee, by, now, dueAt, timely, proof, method]]);
  log_('CYCLE_COMPLETED', cid, tid, by, 'Proof=' + (proof || '(none)') + '; ' + timely + '; identity=' + method +
    (email && method !== 'GOOGLE_EMAIL' ? '; editor email was available but did not match assignee email' : ''));
  if (hadAlert) log_('CYCLE_RESOLVED', cid, tid, 'SYSTEM', 'Alert resolved by completion; further escalation cancelled.');
  var succ = null;
  if (task && task.mode === 'ROLLING' && (task.active || isDemoId_(tid))) {
    succ = createRollingNext_(task, cfg, cid, Number(g('Cycle_No')) || 1, now, now, loadIndex_(), '');
    if (!succ) log_('DUPLICATE_PREVENTED', cid, tid, 'SYSTEM', 'Next rolling cycle already exists or an open cycle exists; not created again.');
  }
  return {ok: true, msg: 'Completed (' + timely + ').' + (succ ? ' Next cycle ' + succ + ' created.' : ''), successor: succ, method: method};
}

// ---------------------------------------------------------------- manager acknowledgement
function handleAck_(sh, row, now, ctx, ticked) {
  var cfg = ctx.cfg, r = sh.getRange(row, 1, 1, CY_H.length).getValues()[0];
  function g(k) { return r[CY[k] - 1]; }
  var cid = g('Cycle_ID'), tid = g('Task_ID');
  if (!cid) return null;
  var cell = sh.getRange(row, CY.Manager_Acknowledged), already = !!g('Acknowledged_At');
  if (!ticked) {
    if (already) { cell.setValue(true); log_('ACK_REJECTED', cid, tid, cfg.managerName, 'Acknowledgement is locked once recorded.'); return {ok: false, msg: 'Acknowledgement is locked once recorded.'}; }
    return null;
  }
  if (already) return {ok: true, dup: true};
  function reject(why) { cell.setValue(false); log_('ACK_REJECTED', cid, tid, cfg.managerName, why); return {ok: false, msg: 'Not recorded: ' + why}; }
  if (g('Status') === 'Completed') return reject('Task is already completed; nothing to acknowledge.');
  if (g('Status') !== 'Overdue' && !g('Manager_Alert_Sent_At')) return reject('No manager alert has been raised for this cycle yet.');
  var email = String(ctx.email || '').toLowerCase(), method, by;
  if (email && cfg.managerEmail && email === cfg.managerEmail.toLowerCase()) { method = 'GOOGLE_EMAIL'; by = email; }
  else { method = 'MANAGER_FALLBACK'; by = cfg.managerName + ' (MANAGER_FALLBACK - not verified)'; }
  sh.getRange(row, CY.Acknowledged_By, 1, 2).setValues([[by, now]]);
  sh.getRange(row, CY.Alert_Status).setValue(g('Leadership_Escalated_At') ? 'Escalated' : 'Acknowledged');
  log_('MANAGER_ACKNOWLEDGED', cid, tid, by, 'Identity method=' + method + '. Acknowledgement does not complete the task.');
  return {ok: true, msg: 'Acknowledgement recorded.', method: method};
}

// ---------------------------------------------------------------- trigger: onEdit (installable)
function onEditHandler(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== SN.CY) return;
  var col = e.range.getColumn();
  if (e.range.getNumColumns() !== 1 || (col !== CY.Done && col !== CY.Manager_Acknowledged)) return;
  var row0 = e.range.getRow(), nrows = e.range.getNumRows();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) { toast_('System busy - please untick and tick again in a minute.'); return; }
  try {
    var cfg = getCfg_(), now = getNow_(cfg), ctx = {cfg: cfg, email: activeEmail_(e)};
    for (var i = 0; i < nrows; i++) {
      var row = row0 + i; if (row < 2) continue;
      var ticked = sh.getRange(row, col).getValue() === true;
      var res = (col === CY.Done) ? handleDone_(sh, row, now, ctx, ticked) : handleAck_(sh, row, now, ctx, ticked);
      if (res && res.msg) toast_(res.msg);
    }
  } catch (err) {
    log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'onEdit: ' + err); toast_('Error: ' + err);
  } finally { lock.releaseLock(); flushLog_(); }
}

// ---------------------------------------------------------------- email
function sendMail_(cfg, to, subject, body, demo) {
  var mode = cfg.mode;
  if (mode !== 'TEST' && mode !== 'LIVE') return {sent: false, reason: 'EMAIL_MODE=' + mode + ' (no emails sent)'};
  var dest = to, prefix = '';
  if (mode === 'TEST' || demo) { dest = cfg.testEmail; prefix = '[TEST] '; body = '[TEST - intended recipient: ' + to + ']\n\n' + body; }
  if (!dest) return {sent: false, reason: 'No recipient configured'};
  try {
    MailApp.sendEmail({to: dest, subject: prefix + subject, body: body, name: ' Recurring Task Tracker'});
    return {sent: true, to: dest};
  } catch (err) { return {sent: false, reason: String(err)}; }
}
function sheetUrl_() { try { return ss_().getUrl(); } catch (e) { return ''; } }

function managerBody_(cfg, items) {
  var lines = items.map(function (it, i) {
    var r = it.r;
    return (i + 1) + ') Employee: ' + r[CY.Assigned_To - 1] + '\n   Task: ' + r[CY.Task_Name - 1] + '\n   Cycle ID: ' + r[CY.Cycle_ID - 1] +
      '\n   Due: ' + it.dueDate + ' ' + ts_(r[CY.Due_Time - 1], cfg.defaultDue) + '\n   Status: Overdue\n   Required proof: ' + r[CY.Required_Proof - 1];
  });
  return 'Hello ' + cfg.managerName + ',\n\nThe following recurring tasks passed their deadline and are not completed:\n\n' + lines.join('\n\n') +
    '\n\nMANAGER ACTION REQUIRED:\n1. Open the sheet: ' + sheetUrl_() + '\n2. Go to the Cycles tab and find the Cycle ID(s) above.\n3. Tick Manager_Acknowledged (optionally add a Manager_Note).\n' +
    'Acknowledging does not complete the task. If a cycle is not acknowledged by ' + cfg.ackTime + ' on the next working day, or is still unresolved by ' + cfg.resTime +
    ' on the next working day, it is escalated to leadership.';
}
function leadershipBody_(cfg, items) {
  var lines = items.map(function (it, i) {
    var r = it.r, ack = r[CY.Manager_Acknowledged - 1] === true;
    return (i + 1) + ') Task: ' + r[CY.Task_Name - 1] + '\n   Employee: ' + r[CY.Assigned_To - 1] + '\n   Cycle ID: ' + r[CY.Cycle_ID - 1] +
      '\n   Original due: ' + it.dueDate + ' ' + ts_(r[CY.Due_Time - 1], cfg.defaultDue) + '\n   Manager alert sent: ' + ds_(r[CY.Manager_Alert_Sent_At - 1] instanceof Date ? r[CY.Manager_Alert_Sent_At - 1] : '') +
      (r[CY.Manager_Alert_Sent_At - 1] instanceof Date ? ' ' + fmt_(r[CY.Manager_Alert_Sent_At - 1], 'HH:mm') : '') +
      '\n   Manager acknowledged: ' + (ack ? 'Yes (' + r[CY.Acknowledged_By - 1] + ')' : 'No') + '\n   Escalation reason: ' + it.reason;
  });
  return 'Leadership escalation - unresolved recurring task(s):\n\n' + lines.join('\n\n') + '\n\nSheet: ' + sheetUrl_() + '\nThis is the only escalation email that will be sent for each listed cycle.';
}

// ---------------------------------------------------------------- the engine
function engineCore_(opts) {
  var cfg = getCfg_(), now = opts.now || getNow_(cfg), demo = !!opts.demo;
  var runId = 'RUN-' + fmt_(new Date(), 'yyyyMMdd-HHmmss') + (demo ? '-DEMO' : '');
  var st = {created: 0, statusChanged: 0, mgrEmails: 0, mgrItems: 0, leadEmails: 0, leadItems: 0, dup: 0, errors: 0, note: ''};
  var sh = sh_(SN.CY);
  if (!demo) { try { generateCycles_(cfg, now, st, runId); } catch (e) { st.errors++; log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'generate: ' + e, runId); } }

  // 1) refresh statuses, collect overdue open cycles
  var items = [], tasks = taskMap_(), n = sh.getLastRow() - 1;
  if (n >= 1) {
    sh.getRange(2, 1, n, CY_H.length).getValues().forEach(function (r, i) {
      var tid = String(r[CY.Task_ID - 1]);
      if (!r[CY.Cycle_ID - 1] || isDemoId_(tid) !== demo) return;
      if (opts.prefix && tid.indexOf(opts.prefix) !== 0) return;
      if (r[CY.Status - 1] === 'Completed') return;
      var t = tasks[tid]; if (!demo && t && !t.active) return;
      var rowNo = i + 2, dueDate = ds_(r[CY.Due_Date - 1]), dueAt = parseAt_(dueDate, ts_(r[CY.Due_Time - 1], cfg.defaultDue));
      var ns = computeStatus_(dueDate, dueAt, now);
      if (ns !== r[CY.Status - 1]) { sh.getRange(rowNo, CY.Status).setValue(ns); r[CY.Status - 1] = ns; st.statusChanged++; }
      if (ns === 'Overdue') items.push({row: rowNo, r: r, dueDate: dueDate, dueAt: dueAt});
    });
  }

  // 2) L1 manager alert (one grouped digest per run)
  var l1 = [];
  items.forEach(function (it) {
    if (it.r[CY.Manager_Alert_Sent_At - 1]) st.dup++;
    else if (now.getTime() >= it.dueAt.getTime() + cfg.delay * 60000) l1.push(it);
  });
  if (l1.length) {
    var res1 = sendMail_(cfg, cfg.managerEmail, 'Action needed: ' + l1.length + ' overdue recurring task(s)', managerBody_(cfg, l1), demo);
    if (res1.sent) {
      st.mgrEmails++; st.mgrItems = l1.length;
      l1.forEach(function (it) {
        sh.getRange(it.row, CY.Manager_Alert_Sent_At).setValue(now);
        sh.getRange(it.row, CY.Alert_Status).setValue(it.r[CY.Manager_Acknowledged - 1] === true ? 'Acknowledged' : 'Manager Alerted');
        it.r[CY.Manager_Alert_Sent_At - 1] = now;
        log_('MANAGER_EMAIL_SENT', it.r[CY.Cycle_ID - 1], it.r[CY.Task_ID - 1], 'SYSTEM', 'Digest of ' + l1.length + ' item(s) sent to ' + res1.to + ' (mode ' + cfg.mode + ')', runId);
      });
    } else {
      st.note = res1.reason;
      if (cfg.mode !== 'OFF') { st.errors++; log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'Manager email not sent: ' + res1.reason, runId); }
    }
  }

  // 3) L2 leadership escalation (max one per cycle)
  var l2 = [];
  items.forEach(function (it) {
    if (it.r[CY.Leadership_Escalated_At - 1]) { st.dup++; return; }
    if (!it.r[CY.Manager_Alert_Sent_At - 1]) return; // manager must be alerted first
    var ack = it.r[CY.Manager_Acknowledged - 1] === true, nwd = nextWorking_(it.dueDate, cfg);
    var ackAt = parseAt_(nwd, cfg.ackTime), resAt = parseAt_(nwd, cfg.resTime), why = '';
    if (!ack && now.getTime() >= ackAt.getTime()) why = 'A: Manager did not acknowledge by ' + nwd + ' ' + cfg.ackTime;
    else if (ack && now.getTime() >= resAt.getTime()) why = 'B: Acknowledged, but task still unresolved at ' + nwd + ' ' + cfg.resTime;
    if (why) { it.reason = why; l2.push(it); }
  });
  if (l2.length) {
    var res2 = sendMail_(cfg, cfg.leadEmail, 'ESCALATION: ' + l2.length + ' unresolved recurring task(s)', leadershipBody_(cfg, l2), demo);
    if (res2.sent) {
      st.leadEmails++; st.leadItems = l2.length;
      l2.forEach(function (it) {
        sh.getRange(it.row, CY.Leadership_Escalated_At).setValue(now);
        sh.getRange(it.row, CY.Escalation_Reason).setValue(it.reason);
        sh.getRange(it.row, CY.Alert_Status).setValue('Escalated');
        it.r[CY.Leadership_Escalated_At - 1] = now;
        log_('LEADERSHIP_EMAIL_SENT', it.r[CY.Cycle_ID - 1], it.r[CY.Task_ID - 1], 'SYSTEM', it.reason + ' -> ' + res2.to + ' (mode ' + cfg.mode + ')', runId);
      });
    } else {
      st.note = res2.reason;
      if (cfg.mode !== 'OFF') { st.errors++; log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'Leadership email not sent: ' + res2.reason, runId); }
    }
  }

  var acted = st.created + st.statusChanged + st.mgrEmails + st.leadEmails + st.errors;
  if (opts.verbose || acted > 0) {
    log_('ENGINE_RUN', '', '', 'SYSTEM', 'effective_now=' + fmt_(now, 'yyyy-MM-dd HH:mm') + (demo ? ' (DEMO)' : '') + '; cycles_created=' + st.created + '; statuses_changed=' + st.statusChanged +
      '; manager_emails=' + st.mgrEmails + '; leadership_emails=' + st.leadEmails + '; already_handled_skipped=' + st.dup + '; errors=' + st.errors + (st.note ? '; note=' + st.note : ''), runId);
    if (opts.verbose && st.dup > 0) log_('DUPLICATE_PREVENTED', '', '', 'SYSTEM', st.dup + ' item(s) already had a cycle/alert/escalation; nothing repeated.', runId);
  }
  return st;
}

// scheduled entry point (time-driven trigger)
function engineTrigger() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;
  try { engineCore_({verbose: false}); }
  catch (err) { log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'engine: ' + err); }
  finally { lock.releaseLock(); flushLog_(); }
}
// menu entry point
function runEngineNow() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) { SpreadsheetApp.getUi().alert('The engine is already running. Try again shortly.'); return; }
  try {
    var st = engineCore_({verbose: true});
    SpreadsheetApp.getUi().alert('Engine run complete',
      'New cycles: ' + st.created + '\nStatus changes: ' + st.statusChanged + '\nManager emails sent: ' + st.mgrEmails + ' (' + st.mgrItems + ' items)\nLeadership emails sent: ' + st.leadEmails + ' (' + st.leadItems +
      ' items)\nAlready handled (not repeated): ' + st.dup + '\nErrors: ' + st.errors + (st.note ? '\nNote: ' + st.note : ''), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) { log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'engine: ' + err); SpreadsheetApp.getUi().alert('Error: ' + err); }
  finally { lock.releaseLock(); flushLog_(); }
}

// ---------------------------------------------------------------- menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Task System')
    .addItem('Setup / Initialise System', 'setupSystem')
    .addItem('Run Engine Now', 'runEngineNow')
    .addItem('Run Demo / Test', 'runDemoTests')
    .addToUi();
}

// ---------------------------------------------------------------- SETUP (non-destructive)
function setupSystem() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  var order = [SN.CY, SN.DB, SN.TM, SN.CF, SN.CP, SN.AL, SN.TC], sheets = {};
  order.forEach(function (n) { sheets[n] = ss.getSheetByName(n) || ss.insertSheet(n); });
  var d = ss.getSheetByName('Sheet1');
  if (d && d.getLastRow() === 0 && ss.getSheets().length > order.length) ss.deleteSheet(d);
  order.forEach(function (n, i) { ss.setActiveSheet(sheets[n]); ss.moveActiveSheet(i + 1); });

  setupConfig_(sheets[SN.CF]);
  setupTaskMaster_(sheets[SN.TM]);
  setupCycles_(sheets[SN.CY]);
  setupLedger_(sheets[SN.CP], CP_H, [8, 9]);
  setupLedger_(sheets[SN.AL], AL_H, [2]);
  setupTests_(sheets[SN.TC]);
  setupDashboard_(sheets[SN.DB]);
  var cfg = getCfg_();
  installTriggers_(cfg);
  ss.setActiveSheet(sheets[SN.CY]);
  log_('SETUP', '', '', 'SYSTEM', 'Setup/initialise executed. Existing records untouched. Triggers reinstalled.');
  flushLog_();
  Logger.log('Setup complete. Sheets, formatting, validations, dashboard and triggers are ready.');
  Logger.log('NEXT: Configure TEST_EMAIL, MANAGER_EMAIL and LEADERSHIP_EMAIL in Config. EMAIL_MODE remains OFF until you deliberately enable testing.');
  Logger.log('Then open the spreadsheet and use Task System > Run Engine Now.');
}

function head_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9d9d9').setWrap(true);
  sh.setFrozenRows(1);
}
function setupConfig_(sh) {
  head_(sh, ['Key', 'Value', 'Notes']);
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  var have = {}, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { have[r[0]] = i + 2; });
  CFG_DEFAULTS.forEach(function (d) {
    if (!have[d[0]]) {
      var v = d[1];
      var r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, 3).setValues([[d[0], v, d[2]]]); have[d[0]] = r;
    }
  });
  sh.getRange(have.EMAIL_MODE, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TEST', 'LIVE', 'OFF'], true).build());
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 520);
}
function setupTaskMaster_(sh) {
  head_(sh, TM_H);
  sh.getRange(2, TM.Due_Time, MAXROWS - 1, 1).setNumberFormat('@');
  sh.getRange(2, TM.Start_Date, MAXROWS - 1, 1).setNumberFormat('@');
  sh.getRange(2, TM.Schedule_Mode, MAXROWS - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CALENDAR_DAILY', 'ROLLING'], true).setAllowInvalid(false).build());
  sh.getRange(2, TM.Proof_Type, MAXROWS - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['WHATSAPP_CHECKBOX', 'TICKET_ID'], true).setAllowInvalid(false).build());
  sh.getRange(2, TM.Active, MAXROWS - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  if (sh.getLastRow() <= 1) {
    var today = fmt_(new Date(), 'yyyy-MM-dd');
    appendRows_(sh, [
      ['T001', 'Daily Operations Update', 'Employee A', 'employee.a@example.com', 1, 'CALENDAR_DAILY', 'WHATSAPP_CHECKBOX', '17:00', today, true, 'Share update in the WhatsApp group, then tick Done in Cycles'],
      ['T002', 'Portal Task A', 'Employee B', 'employee.b@example.com', 7, 'ROLLING', 'TICKET_ID', '17:00', today, true, 'Every 7 days after completion'],
      ['T003', 'Portal Task B', 'Employee C', 'employee.c@example.com', 15, 'ROLLING', 'TICKET_ID', '17:00', today, true, 'Every 15 days after completion'],
      ['T004', 'Portal Task C', 'Employee D', 'employee.d@example.com', 30, 'ROLLING', 'TICKET_ID', '17:00', today, true, 'Every 30 days after completion']
    ]);
  }
  sh.getRange(1, 1, 1, TM_H.length).setBackground('#cfe2f3');
  [90, 200, 130, 210, 110, 140, 160, 80, 100, 60, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}
function setupCycles_(sh) {
  head_(sh, CY_H);
  var hc = sh.getRange(1, 1, 1, CY_H.length);
  hc.setBackground('#d9d9d9');
  sh.getRange(1, 1, 1, 9).setBackground('#cfe2f3');                 // employee-facing
  sh.getRange(1, CY.Ticket_ID_or_Proof, 1, 2).setBackground('#ffe599'); // employee input
  sh.getRange(1, CY.Completed_By, 1, 4).setBackground('#b6d7a8');   // completion tracking
  sh.getRange(1, CY.Manager_Alert_Sent_At, 1, 8).setBackground('#ead1dc'); // manager/escalation (system)
  sh.getRange(1, CY.Manager_Acknowledged).setBackground('#f9cb9c');  // manager input
  sh.getRange(1, CY.Manager_Note).setBackground('#f9cb9c');
  sh.getRange(1, CY.Cycle_No, 1, 5).setBackground('#d9d9d9');       // system
  [CY.Due_Date, CY.Due_Time, CY.Schedule_Date].forEach(function (c) { sh.getRange(2, c, MAXROWS - 1, 1).setNumberFormat('@'); });
  [CY.Completed_At, CY.Manager_Alert_Sent_At, CY.Acknowledged_At, CY.Leadership_Escalated_At, CY.Due_At, CY.Created_At].forEach(function (c) {
    sh.getRange(2, c, MAXROWS - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  });
  var cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, CY.Done, MAXROWS - 1, 1).setDataValidation(cb);
  sh.getRange(2, CY.Manager_Acknowledged, MAXROWS - 1, 1).setDataValidation(cb);
  sh.getRange(2, CY.Ticket_ID_or_Proof, MAXROWS - 1, 2).setBackground('#fff2cc');
  sh.getRange(2, CY.Manager_Acknowledged, MAXROWS - 1, 1).setBackground('#fce5cd');
  sh.getRange(2, CY.Manager_Note, MAXROWS - 1, 1).setBackground('#fce5cd');
  var rng = sh.getRange(2, CY.Status, MAXROWS - 1, 1), rules = [];
  [['Overdue', '#f4c7c3'], ['Due', '#fce8b2'], ['Completed', '#b7e1cd'], ['Scheduled', '#e0e0e0']].forEach(function (p) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(p[0]).setBackground(p[1]).setRanges([rng]).build());
  });
  sh.setConditionalFormatRules(rules);
  // warning-only protection on system-managed columns (script keeps working; humans get a warning)
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { if (p.getDescription() === 'SYSTEM_MANAGED') p.remove(); });
  [[1, 9], [CY.Completed_By, 4], [CY.Manager_Alert_Sent_At, 1], [CY.Acknowledged_By, 2], [CY.Leadership_Escalated_At, 3], [CY.Cycle_No, 5]].forEach(function (g) {
    sh.getRange(2, g[0], MAXROWS - 1, g[1] === 9 && g[0] === 1 ? 9 : g[1]).protect().setDescription('SYSTEM_MANAGED').setWarningOnly(true);
  });
  sh.setFrozenColumns(1);
  [170, 80, 190, 120, 90, 70, 90, 140, 300, 160, 55, 190, 130, 150, 90, 130, 110, 190, 130, 200, 130, 280, 120, 70, 100, 130, 130, 170].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}
function setupLedger_(sh, headers, dateCols) {
  head_(sh, headers);
  dateCols.forEach(function (c) { sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); });
  for (var i = 1; i <= headers.length; i++) sh.setColumnWidth(i, i === 7 && headers === AL_H ? 520 : 140);
}
function setupTests_(sh) {
  head_(sh, TC_H);
  if (sh.getLastRow() > 1) return; // keep any recorded results
  var t = [
    ['T1', 'Recurring cycle (7-day task)', 'Complete a 7-day ticket task (Ticket ID + Done)', 'Old cycle closes; exactly one next cycle created, due = completion date + 7 days (Sunday/holiday moves to next working day); early completion of the new cycle is blocked'],
    ['T2', 'Single-point updating (15-day task)', 'Employee enters Ticket ID and ticks Done in Cycles only', 'Blank Ticket ID rejected; with Ticket ID the cycle completes and Completions + Audit_Log rows appear with no manual update elsewhere'],
    ['T3', 'Daily task via WhatsApp checkbox', 'Tick Done on a daily cycle with no Ticket ID; a missed earlier daily cycle exists', 'Completes without Ticket ID; no successor generated by completion; missed earlier cycle stays Overdue and independent; duplicate daily cycle blocked'],
    ['T4', 'Same-day manager alert', 'Past-due incomplete cycles; run engine after deadline + 30 min', 'One grouped manager email (to TEST_EMAIL in TEST mode); Manager_Alert_Sent_At set'],
    ['T5', 'Manager acknowledgement', 'Tick Manager_Acknowledged on an alerted overdue cycle', 'Acknowledged_By, Acknowledged_At and Alert_Status recorded; task stays Overdue (acknowledgement does not complete it)'],
    ['T6', 'Leadership escalation (both conditions)', 'A: unacknowledged past ack deadline. B: acknowledged but unresolved past resolution deadline', 'One leadership email per cycle with reason A / reason B; Leadership_Escalated_At set'],
    ['T7', 'Escalation timing thresholds', 'Engine run just before and just after each threshold (17:29/17:31, next WD 10:29/10:31, 16:59/17:01)', 'No action before a threshold; correct action immediately after it'],
    ['T8', 'No duplicates on repeated runs', 'Run engine twice more with no data changes; retry cycle creation', 'No new cycles, completions, manager or leadership emails; duplicates logged as prevented']
  ];
  appendRows_(sh, t.map(function (r) { return [r[0], r[1], r[2], r[3], 'NOT RUN', 'NOT RUN', 'Not executed yet. Use Task System > Run Demo / Test, plus the manual live checks in Documentation.md']; }));
  sh.getRange(2, 1, 8, 7).setWrap(true).setVerticalAlignment('top');
  [60, 220, 320, 380, 360, 90, 420].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}
function setupDashboard_(sh) {
  sh.clear();
  var ex = '"<>' + DEMO_PREFIX + '*"', B = 'Cycles!$B$2:$B', G = 'Cycles!$G$2:$G';
  function cnt(status) { return '=COUNTIFS(' + G + ',"' + status + '",' + B + ',' + ex + ')'; }
  sh.getRange('A1').setValue('MANAGER DASHBOARD - recurring tasks').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Formula-driven. Statuses refresh on every engine run (every 15 minutes). DEMO_ test records are excluded.').setFontStyle('italic');
  var kpi = [
    ['Total active tasks', '=COUNTIFS(Task_Master!$J$2:$J,TRUE,Task_Master!$A$2:$A,' + ex + ')'],
    ['Scheduled (future)', cnt('Scheduled')],
    ['Currently due (not yet overdue)', cnt('Due')],
    ['Overdue', cnt('Overdue')],
    ['Pending (due + overdue)', '=B6+B7'],
    ['Completed (all cycles)', cnt('Completed')],
    ['Escalated to leadership', '=COUNTIFS(Cycles!$U$2:$U,"<>",' + B + ',' + ex + ')'],
    ['Awaiting manager acknowledgement', '=COUNTIFS(Cycles!$P$2:$P,"<>",Cycles!$Q$2:$Q,"<>TRUE",' + G + ',"<>Completed",' + B + ',' + ex + ')']
  ];
  sh.getRange('A4:B4').setValues([['Measure', 'Count']]).setFontWeight('bold').setBackground('#d9d9d9');
  sh.getRange(5, 1, kpi.length, 2).setValues(kpi);
  // rows: 5 total, 6 scheduled, 7 due, 8 overdue, 9 pending ... fix pending formula to correct rows
  sh.getRange('B9').setFormula('=B7+B8');
  sh.getRange('A15:D15').setValues([['Employee', 'Due', 'Overdue', 'Open total']]).setFontWeight('bold').setBackground('#d9d9d9');
  sh.getRange('A14').setValue('Employee-wise pending').setFontWeight('bold');
  for (var i = 16; i <= 25; i++) {
    sh.getRange(i, 1).setFormula('=IFERROR(INDEX(SORT(UNIQUE(FILTER(Task_Master!$C$2:$C,Task_Master!$C$2:$C<>"",NOT(REGEXMATCH(Task_Master!$A$2:$A,"^DEMO_"))))),ROW()-15),"")');
    sh.getRange(i, 2).setFormula('=IF($A' + i + '="","",COUNTIFS(Cycles!$D$2:$D,$A' + i + ',' + G + ',"Due",' + B + ',' + ex + '))');
    sh.getRange(i, 3).setFormula('=IF($A' + i + '="","",COUNTIFS(Cycles!$D$2:$D,$A' + i + ',' + G + ',"Overdue",' + B + ',' + ex + '))');
    sh.getRange(i, 4).setFormula('=IF($A' + i + '="","",B' + i + '+C' + i + ')');
  }
  sh.getRange('A27').setValue('Task-wise status overview').setFontWeight('bold');
  sh.getRange('A28:F28').setValues([['Task', 'Scheduled', 'Due', 'Overdue', 'Completed', 'Escalated']]).setFontWeight('bold').setBackground('#d9d9d9');
  for (var j = 29; j <= 38; j++) {
    sh.getRange(j, 1).setFormula('=IFERROR(INDEX(SORT(UNIQUE(FILTER(Task_Master!$B$2:$B,Task_Master!$B$2:$B<>"",NOT(REGEXMATCH(Task_Master!$A$2:$A,"^DEMO_"))))),ROW()-28),"")');
    ['Scheduled', 'Due', 'Overdue', 'Completed'].forEach(function (s, k) {
      sh.getRange(j, 2 + k).setFormula('=IF($A' + j + '="","",COUNTIFS(Cycles!$C$2:$C,$A' + j + ',' + G + ',"' + s + '",' + B + ',' + ex + '))');
    });
    sh.getRange(j, 6).setFormula('=IF($A' + j + '="","",COUNTIFS(Cycles!$C$2:$C,$A' + j + ',Cycles!$U$2:$U,"<>",' + B + ',' + ex + '))');
  }
  sh.getRange('H3').setValue('Outstanding items (Due + Overdue)').setFontWeight('bold');
  sh.getRange('H4:M4').setValues([['Cycle_ID', 'Employee', 'Task', 'Due_Date', 'Status', 'Alert_Status']]).setFontWeight('bold').setBackground('#d9d9d9');
  sh.getRange('H5').setFormula('=IFERROR(QUERY(Cycles!$A$2:$W,"select A,D,C,E,G,W where (G=\'Overdue\' or G=\'Due\') and not B starts with \'DEMO_\' order by E",0),"No outstanding tasks")');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 90); sh.setColumnWidth(3, 90); sh.setColumnWidth(4, 90);
  [170, 130, 200, 100, 90, 130].forEach(function (w, i) { sh.setColumnWidth(8 + i, w); });
}
function installTriggers_(cfg) {
  var mine = {onEditHandler: 1, engineTrigger: 1};
  ScriptApp.getProjectTriggers().forEach(function (t) { if (mine[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss_()).onEdit().create();
  var m = [1, 5, 10, 15, 30].indexOf(cfg.interval) > -1 ? cfg.interval : 15;
  ScriptApp.newTrigger('engineTrigger').timeBased().everyMinutes(m).create();
}

// ---------------------------------------------------------------- DEMO / TEST RUNNER (isolated, additive only)
function updateTest_(id, pass, actual, notes) {
  var sh = sh_(SN.TC), n = sh.getLastRow() - 1; if (n < 1) return;
  var ids = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) {
    if (ids[i][0] === id) { sh.getRange(i + 2, 5, 1, 3).setValues([[actual, pass ? 'PASS' : 'FAIL', notes]]); return; }
  }
}

function runDemoTests() {
  var ui = SpreadsheetApp.getUi(), lock = LockService.getScriptLock();
  var cfg = getCfg_();
  if (cfg.mode !== 'TEST' && cfg.mode !== 'LIVE') { ui.alert('Set EMAIL_MODE to TEST in Config first (demo emails always go to TEST_EMAIL).'); return; }
  if (!cfg.testEmail) { ui.alert('Set TEST_EMAIL in Config first.'); return; }
  if (ui.alert('Run demo / test', 'This ADDS isolated DEMO_ records (Active = FALSE, excluded from the dashboard and real engine) and sends about 3 test emails to ' + cfg.testEmail +
    '. Nothing existing is erased or overwritten, except the Actual_Result/Status cells in Test_Cases. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  if (!lock.tryLock(30000)) { ui.alert('System busy. Try again shortly.'); return; }
  var summary = '';
  try {
    var out = demoCore_(cfg);
    summary = Object.keys(out).sort().map(function (k) { return k + ': ' + (out[k].pass ? 'PASS' : 'FAIL') + ' - ' + out[k].actual; }).join('\n\n');
  } catch (err) {
    log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'demo: ' + err); summary = 'Demo aborted: ' + err;
  } finally { lock.releaseLock(); flushLog_(); }
  ui.alert('Demo / test results (also written to Test_Cases)', summary.substr(0, 1800), ui.ButtonSet.OK);
}

function demoCore_(cfg) {
  var cy = sh_(SN.CY), tm = sh_(SN.TM), cp = sh_(SN.CP);
  var real = new Date(), today = fmt_(real, 'yyyy-MM-dd');
  var base = prevWorking_(today, cfg), prevDay = prevWorking_(addDays_(base, -1), cfg), nextWD = nextWorking_(base, cfg);
  var P = DEMO_PREFIX + fmt_(real, 'MMddHHmmss') + '_' + Utilities.getUuid().substring(0, 6) + '_';
  var runNote = 'Executed by Run Demo/Test ' + fmt_(real, 'yyyy-MM-dd HH:mm') + ' (IST), run key ' + P + ', simulated base date ' + base +
    '. Handlers were called directly (not through a live checkbox click / time trigger). ';
  var idn = 'demo.employee@example.com';
  var defs = [['D7', 'Demo 7-day Portal Task', 'Demo Employee B', 7, 'ROLLING', 'TICKET_ID'], ['D15', 'Demo 15-day Portal Task', 'Demo Employee C', 15, 'ROLLING', 'TICKET_ID'],
    ['DD', 'Demo Daily Update', 'Demo Employee A', 1, 'CALENDAR_DAILY', 'WHATSAPP_CHECKBOX'], ['EA', 'Demo Escalation A (unacknowledged)', 'Demo Employee C', 7, 'ROLLING', 'TICKET_ID'],
    ['EB', 'Demo Escalation B (acknowledged)', 'Demo Employee D', 7, 'ROLLING', 'TICKET_ID']];
  var T = {}, rows = [];
  defs.forEach(function (d) {
    var id = P + d[0];
    T[d[0]] = {id: id, name: d[1], who: d[2], email: idn, freq: d[3], mode: d[4], proof: d[5], due: '17:00', start: base, active: false};
    rows.push([id, d[1], d[2], idn, d[3], d[4], d[5], '17:00', base, false, 'DEMO - isolated test record (Active=FALSE; ignored by real engine and dashboard)']);
  });
  appendRows_(tm, rows);

  function rowOf(id) { var n = cy.getLastRow() - 1, v = cy.getRange(2, 1, n, 1).getValues(); for (var i = 0; i < n; i++) if (v[i][0] === id) return i + 2; return 0; }
  function cell(id, k) { return cy.getRange(rowOf(id), CY[k]).getValue(); }
  function mkRolling(key, due, no, nowS) {
    var t = T[key], id = t.id + '-C' + pad_(no, 3);
    appendRows_(cy, [cycleRow_(t, cfg, id, due, no, due, '', nowS)]);
    log_('CYCLE_CREATED', id, t.id, 'DEMO', 'Demo cycle created');
    return id;
  }
  function tick(id, proof, nowS) {
    var r = rowOf(id);
    if (proof !== undefined) cy.getRange(r, CY.Ticket_ID_or_Proof).setValue(proof);
    cy.getRange(r, CY.Done).setValue(true);
    return handleDone_(cy, r, nowS, {cfg: cfg, email: ''}, true);
  }
  function ledgerRows(cid) {
    var n = cp.getLastRow() - 1; if (n < 1) return [];
    return cp.getRange(2, 1, n, CP_H.length).getValues().filter(function (r) { return r[CP.Cycle_ID - 1] === cid; });
  }
  var out = {}, S = {};
  function run(id, fn) {
    var r;
    try { r = fn(); } catch (err) { r = {pass: false, actual: 'Exception: ' + err}; log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'demo ' + id + ': ' + err); }
    out[id] = r; updateTest_(id, r.pass, r.actual, runNote + (r.note || ''));
  }

  // ---- T1 recurring 7-day
  run('T1', function () {
    var nowS = parseAt_(base, '15:00'), c1 = mkRolling('D7', base, 1, nowS); S.c1 = c1;
    var res = tick(c1, 'DEMO-7001', nowS); SpreadsheetApp.flush();
    var all = allCycles_(), succ = all.filter(function (r) { return r[CY.Previous_Cycle_ID - 1] === c1; });
    var expDue = onOrAfterWorking_(addDays_(base, 7), cfg);
    var open = all.filter(function (r) { return r[CY.Task_ID - 1] === T.D7.id && r[CY.Status - 1] !== 'Completed'; }).length;
    var sId = succ.length ? succ[0][CY.Cycle_ID - 1] : '';
    var early = sId ? tick(sId, 'DEMO-7002', nowS) : {ok: true};
    var ok = res.ok && cell(c1, 'Status') === 'Completed' && succ.length === 1 && ds_(succ[0][CY.Due_Date - 1]) === expDue && open === 1;
    var okEarly = !early.ok && cell(sId, 'Status') === 'Scheduled';
    return {pass: ok && okEarly, actual: 'Old cycle ' + c1 + ' = ' + cell(c1, 'Status') + '; successors found: ' + succ.length + ' (' + sId + ' due ' + (succ.length ? ds_(succ[0][CY.Due_Date - 1]) : '-') +
      ', expected ' + expDue + '); open cycles for task: ' + open + '; early completion of successor blocked: ' + okEarly};
  });
  // ---- T2 single-point updating
  run('T2', function () {
    var nowS = parseAt_(base, '15:30'), c = mkRolling('D15', base, 1, nowS), r = rowOf(c);
    cy.getRange(r, CY.Done).setValue(true);
    var rej = handleDone_(cy, r, nowS, {cfg: cfg, email: ''}, true);   // no ticket yet
    var res = tick(c, 'DEMO-1501', nowS); SpreadsheetApp.flush();
    var led = ledgerRows(c), expDue = onOrAfterWorking_(addDays_(base, 15), cfg);
    var succ = allCycles_().filter(function (x) { return x[CY.Previous_Cycle_ID - 1] === c; });
    var ok = !rej.ok && res.ok && led.length === 1 && led[0][CP.Ticket_ID_or_Proof - 1] === 'DEMO-1501' && cell(c, 'Status') === 'Completed' &&
      cell(c, 'Identity_Method') === 'ASSIGNEE_FALLBACK' && succ.length === 1 && ds_(succ[0][CY.Due_Date - 1]) === expDue;
    return {pass: ok, actual: 'Blank Ticket ID rejected: ' + !rej.ok + '; after Ticket ID + Done: status ' + cell(c, 'Status') + ', ledger rows ' + led.length + ' (ticket ' + (led.length ? led[0][CP.Ticket_ID_or_Proof - 1] : '-') +
      '), identity ' + cell(c, 'Identity_Method') + ' (fallback, not verified), next due ' + (succ.length ? ds_(succ[0][CY.Due_Date - 1]) : '-') + ' expected ' + expDue + '. Only the Cycles row was edited by the "employee".'};
  });
  // ---- T3 daily task
  run('T3', function () {
    var nowS = parseAt_(base, '16:00'), ix = loadIndex_();
    var cPrev = createDailyIfMissing_(T.DD, cfg, prevDay, nowS, ix, ''), cToday = createDailyIfMissing_(T.DD, cfg, base, nowS, ix, '');
    var res = tick(cToday, undefined, nowS); SpreadsheetApp.flush();
    var dup = createDailyIfMissing_(T.DD, cfg, base, nowS, loadIndex_(), '');
    var all = allCycles_().filter(function (r) { return r[CY.Task_ID - 1] === T.DD.id; });
    var ok = res.ok && cell(cToday, 'Status') === 'Completed' && ledgerRows(cToday).length === 1 && all.length === 2 && cell(cPrev, 'Status') === 'Overdue' && dup === null;
    return {pass: ok, actual: 'Daily cycle ' + cToday + ' completed with no Ticket ID (identity ' + cell(cToday, 'Identity_Method') + '); ledger rows ' + ledgerRows(cToday).length +
      '; total daily cycles ' + all.length + ' (no successor from completion); missed cycle ' + cPrev + ' still ' + cell(cPrev, 'Status') + '; duplicate daily creation blocked: ' + (dup === null),
      note: 'WhatsApp posting itself cannot be verified - checkbox is self-attestation. '};
  });
  // ---- T4..T8 SLA chain
  var chain = ['T4', 'T5', 'T6', 'T7', 'T8'];
  chain.forEach(function (id) { out[id] = {pass: false, actual: 'Not reached'}; });
  try {
    var nowC = parseAt_(base, '16:00'), eA = mkRolling('EA', base, 1, nowC), eB = mkRolling('EB', base, 1, nowC);
    function eng(d, t) { return engineCore_({demo: true, prefix: P + 'E', verbose: true, now: parseAt_(d, t)}); }
    var s1 = eng(base, '17:29'), s2 = eng(base, '17:31');
    var stampedA = cell(eA, 'Manager_Alert_Sent_At') instanceof Date, stampedB = cell(eB, 'Manager_Alert_Sent_At') instanceof Date;
    out.T4 = {pass: s1.mgrEmails === 0 && s2.mgrEmails === 1 && s2.mgrItems === 2 && stampedA && stampedB,
      actual: 'Engine at 17:29 sent ' + s1.mgrEmails + ' manager email(s); at 17:31 sent ' + s2.mgrEmails + ' digest with ' + s2.mgrItems + ' items to ' + cfg.testEmail + '; Manager_Alert_Sent_At set on both: ' + (stampedA && stampedB) + (s2.note ? ' (' + s2.note + ')' : ''),
      note: 'Check the inbox of TEST_EMAIL for subject "[TEST] Action needed...". '};
    // T5 acknowledge EB
    var rB = rowOf(eB); cy.getRange(rB, CY.Manager_Acknowledged).setValue(true);
    var ack = handleAck_(cy, rB, parseAt_(base, '17:40'), {cfg: cfg, email: ''}, true); SpreadsheetApp.flush();
    out.T5 = {pass: ack && ack.ok && cell(eB, 'Manager_Acknowledged') === true && cell(eB, 'Acknowledged_At') instanceof Date && String(cell(eB, 'Acknowledged_By')).indexOf('MANAGER_FALLBACK') > -1 &&
      cell(eB, 'Alert_Status') === 'Acknowledged' && cell(eB, 'Status') === 'Overdue',
      actual: 'Acknowledged_By = "' + cell(eB, 'Acknowledged_By') + '"; Alert_Status = ' + cell(eB, 'Alert_Status') + '; underlying Status still ' + cell(eB, 'Status') + ' (ack does not complete the task)',
      note: 'Identity was unavailable in the demo so MANAGER_FALLBACK was recorded. Live click test in Documentation.md. '};
    // escalation timeline
    var s3 = eng(nextWD, '10:29'), s4 = eng(nextWD, '10:31');
    var aEsc = cell(eA, 'Leadership_Escalated_At') instanceof Date, bEsc4 = cell(eB, 'Leadership_Escalated_At') instanceof Date, reasonA = cell(eA, 'Escalation_Reason');
    var s5 = eng(nextWD, '16:59'), bEsc5 = cell(eB, 'Leadership_Escalated_At') instanceof Date;
    var s6 = eng(nextWD, '17:01'), bEsc6 = cell(eB, 'Leadership_Escalated_At') instanceof Date, reasonB = cell(eB, 'Escalation_Reason');
    var okA = s3.leadEmails === 0 && s4.leadEmails === 1 && aEsc && !bEsc4 && String(reasonA).charAt(0) === 'A';
    var okB = s5.leadEmails === 0 && !bEsc5 && s6.leadEmails === 1 && bEsc6 && String(reasonB).charAt(0) === 'B';
    out.T6 = {pass: okA && okB, actual: 'Condition A: unacknowledged cycle escalated at ' + nextWD + ' 10:31 (reason "' + reasonA + '"), acknowledged cycle not escalated then. Condition B: acknowledged cycle escalated at 17:01 (reason "' + reasonB + '"). Leadership emails: ' + s4.leadEmails + ' + ' + s6.leadEmails};
    out.T7 = {pass: s1.mgrEmails === 0 && s2.mgrEmails === 1 && s3.leadEmails === 0 && s4.leadEmails === 1 && s5.leadEmails === 0 && s6.leadEmails === 1,
      actual: 'Manager alert: 17:29 -> ' + s1.mgrEmails + ', 17:31 -> ' + s2.mgrEmails + '. Leadership: next WD 10:29 -> ' + s3.leadEmails + ', 10:31 -> ' + s4.leadEmails + ', 16:59 -> ' + s5.leadEmails + ', 17:01 -> ' + s6.leadEmails,
      note: 'Simulated clock passed to the engine; Config DEMO_NOW_OVERRIDE was not touched. '};
    // T8 duplicates
    var cyBefore = allCycles_().length, cpBefore = cp.getLastRow();
    var s7 = eng(nextWD, '17:05'), s8 = eng(nextWD, '17:10');
    var dupDaily = createDailyIfMissing_(T.DD, cfg, base, parseAt_(base, '16:00'), loadIndex_(), '');
    var dupRoll = createRollingNext_(T.D7, cfg, S.c1, 1, parseAt_(base, '15:00'), parseAt_(base, '15:00'), loadIndex_(), '');
    var noChange = allCycles_().length === cyBefore && cp.getLastRow() === cpBefore;
    out.T8 = {pass: s7.mgrEmails === 0 && s8.mgrEmails === 0 && s7.leadEmails === 0 && s8.leadEmails === 0 && s7.created === 0 && s8.created === 0 && noChange && s7.dup > 0 && s8.dup > 0 && dupDaily === null && dupRoll === null,
      actual: 'Two further engine runs: manager emails ' + s7.mgrEmails + '/' + s8.mgrEmails + ', leadership emails ' + s7.leadEmails + '/' + s8.leadEmails + ', new cycles ' + s7.created + '/' + s8.created +
        '; cycle rows and ledger rows unchanged: ' + noChange + '; already-handled items skipped: ' + s7.dup + '/' + s8.dup + '; duplicate daily cycle blocked: ' + (dupDaily === null) + '; duplicate rolling successor blocked: ' + (dupRoll === null)};
  } catch (err) {
    log_('AUTOMATION_ERROR', '', '', 'SYSTEM', 'demo chain: ' + err);
    chain.forEach(function (id) { if (out[id].actual === 'Not reached') out[id] = {pass: false, actual: 'Exception before this check: ' + err}; });
  }
  chain.forEach(function (id) { updateTest_(id, out[id].pass, out[id].actual, runNote + (out[id].note || '')); });
  log_('DEMO_RUN', '', '', 'SYSTEM', 'Demo run ' + P + ': ' + Object.keys(out).sort().map(function (k) { return k + '=' + (out[k].pass ? 'PASS' : 'FAIL'); }).join(', '));
  return out;
}
