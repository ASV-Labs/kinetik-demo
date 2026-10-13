/**
 * G2 PSG Tracker - automation layer (Google Apps Script, bound to the tracker spreadsheet).
 *
 * Install: in the tracker, Extensions > Apps Script, replace Code.gs with this file, save,
 * reload the spreadsheet, then run  G2 PSG > Setup / repair  once and approve the permissions.
 *
 * Everything the DA 6s need is already done by formulas; this script adds the conveniences:
 * menu, auto-fill, BDE roster import, shared view, request form, daily brief, archive/close-out,
 * mind map and chart rebuild.
 */

var TABS = {
  DASH: 'Dashboard', SLOTS: 'Duty Slots', NEXT: 'Next Up', PERS: 'Personnel', AVAIL: 'Availability',
  APPTS: 'Appointments', EQUIP: 'Equipment', TASK: 'Taskings', SET: 'Settings', DATA: 'Dash Data',
  REQ: 'Requests'
};
// one DA 6 per duty (AR 220-45 para 4b); bfCol = that roster's "BF #" column on Personnel (J..O)
var ROSTERS = [
  { key: 'CQ NCO', tab: 'DA6 CQ NCO', bfCol: 10 },
  { key: 'CQ RUNNER', tab: 'DA6 CQ RUNNER', bfCol: 11 },
  { key: 'SD RUNNER', tab: 'DA6 SD RUNNER', bfCol: 12 },
  { key: 'SDNCO', tab: 'DA6 SDNCO', bfCol: 13 },
  { key: 'FLAG', tab: 'DA6 FLAG', bfCol: 14 },
  { key: 'ROAD GUARD', tab: 'DA6 ROAD GUARD', bfCol: 15 }
];
// Duty Slots columns (1-based); data starts on row 4
var S = { DATE: 1, END: 2, DUTY: 4, POS: 5, SOLDIER: 6, STATUS: 7, RANK: 8, SECTION: 9, ROSTER: 10,
          CHECK: 12, SUGGEST: 13, REMARKS: 14, SOURCE: 15 };
var FIRST_ROW = 4;
var RANKS = ['PV1', 'PV2', 'PFC', 'SPC', 'CPL', 'SGT', 'SSG', 'SFC', 'MSG', '1SG', 'SGM', 'CSM',
             'WO1', 'CW2', 'CW3', 'CW4', 'CW5', '2LT', '1LT', 'CPT', 'MAJ', 'LTC', 'COL', 'PVT'];
var OFFICER_RANKS = ['WO1', 'CW2', 'CW3', 'CW4', 'CW5', '2LT', '1LT', 'CPT', 'MAJ', 'LTC', 'COL'];

// ---------------------------------------------------------------------------------------------
// Menu & simple triggers
// ---------------------------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('G2 PSG')
    .addItem('Setup / repair (run once)', 'setup')
    .addSeparator()
    .addItem('Auto-fill open G2 slots', 'autoFillOpenSlots')
    .addItem('Import G2 slots from a BDE roster tab...', 'importFromRosterPrompt')
    .addItem('Sort Duty Slots by date', 'sortSlots')
    .addItem('Mark past slots Completed', 'markPastCompleted')
    .addSeparator()
    .addItem('Publish shared view (no PII)', 'publishSharedView')
    .addItem('Build / refresh one-slider brief', 'buildOneSlider')
    .addItem('Create exemption request form', 'createRequestForm')
    .addItem('Email me the brief now', 'emailBrief')
    .addSeparator()
    .addItem('Add tracker table...', 'addTrackerTablePrompt')
    .addItem('Close & restart DA 6s (archive)...', 'closeRosters')
    .addItem('Rebuild mind map', 'buildMindMap')
    .addItem('Rebuild dashboard charts', 'rebuildCharts')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var name = sh.getName();
  if (name === TABS.SLOTS && e.range.getRow() >= FIRST_ROW) {
    var row = e.range.getRow();
    var col = e.range.getColumn();
    // picking a Soldier on an open slot marks it Scheduled
    if (col === S.SOLDIER && e.value && !sh.getRange(row, S.STATUS).getValue()) {
      sh.getRange(row, S.STATUS).setValue('Scheduled');
    }
  }
  if (name === TABS.REQ && e.range.getRow() >= 2) {
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var approveCol = hdr.indexOf('Approve') + 1;
    if (e.range.getColumn() === approveCol && (e.value === 'TRUE' || e.value === true)) {
      approveRequestRow_(sh, e.range.getRow(), hdr);
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------------------------
function setup() {
  var ss = SpreadsheetApp.getActive();
  applyCosmetics_(ss);
  protectComputed_(ss);
  installTriggers_();
  try { buildMindMap(); } catch (err) { Logger.log('mind map: ' + err); }
  if (ss.getSheetByName(TABS.DASH).getCharts().length < 7) rebuildCharts();
  SpreadsheetApp.getUi().alert('G2 PSG setup complete',
    'Menu, daily 0500 automation, formatting, protections, mind map and charts are in place.\n\n' +
    'Daily job: marks past slots Completed, refreshes the shared view and one-slider (once created) and emails you the brief.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function applyCosmetics_(ss) {
  var fmt = 'dd mmm yy';
  var dateCols = [[TABS.SLOTS, ['A', 'B', 'P']], [TABS.AVAIL, ['C', 'D', 'M']], [TABS.PERS, ['Z', 'AA']],
                  [TABS.APPTS, ['A']], [TABS.EQUIP, ['G', 'J']], [TABS.TASK, ['B', 'C']]];
  dateCols.forEach(function (d) {
    var sh = ss.getSheetByName(d[0]);
    if (!sh) return;
    d[1].forEach(function (c) { sh.getRange(c + '4:' + c).setNumberFormat(fmt); });
  });
  ROSTERS.forEach(function (r) {
    var sh = ss.getSheetByName(r.tab);
    if (!sh) return;
    sh.getRange('I8:CA127').setHorizontalAlignment('center').setFontSize(9)
      .setBorder(null, null, null, null, true, true, '#D5DAE1', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange('A8:H127').setHorizontalAlignment('center').setFontSize(10)
      .setBorder(null, null, null, null, null, true, '#D5DAE1', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange('B8:B127').setHorizontalAlignment('left').setFontWeight('bold');
    sh.getRange('I5:CA7').setHorizontalAlignment('center').setFontSize(8);
    sh.setRowHeights(3, 2, 18);
  });
  var nu = ss.getSheetByName(TABS.NEXT);
  if (nu) nu.getRange('A6:AJ125').setFontSize(9).setHorizontalAlignment('center');
}

function protectComputed_(ss) {
  var spec = [
    [TABS.SLOTS, ['C4:C', 'H4:M', 'P4:P']], [TABS.AVAIL, ['G4:M']], [TABS.PERS, ['P4:AA']],
    [TABS.APPTS, ['H4:K']], [TABS.EQUIP, ['I4:K']], [TABS.TASK, ['L4:M']], [TABS.DATA, ['A1:AH60']]
  ];
  ROSTERS.forEach(function (r) { spec.push([r.tab, ['A1:EU127']]); });
  spec.push([TABS.NEXT, ['A1:AJ125']]);
  spec.push([TABS.DASH, ['A1:N140']]);
  spec.forEach(function (s) {
    var sh = ss.getSheetByName(s[0]);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
      if (p.getDescription() === 'G2 PSG auto') p.remove();
    });
    s[1].forEach(function (a1) {
      sh.getRange(a1).protect().setDescription('G2 PSG auto').setWarningOnly(true);
    });
  });
}

function installTriggers_() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'dailyJob' || f === 'onRequestSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(5).create();
  var formUrl = settingValue_('Exemption request form URL');
  if (formUrl) ScriptApp.newTrigger('onRequestSubmit').forSpreadsheet(ss).onFormSubmit().create();
}

function dailyJob() {
  markPastCompleted(true);
  if (settingValue_('Shared view file ID')) publishSharedView(true);
  if (settingValue_('One-slider file ID')) buildOneSlider(true);
  emailBrief(true);
}

// ---------------------------------------------------------------------------------------------
// Settings helpers (column A label / column B value on the Settings tab)
// ---------------------------------------------------------------------------------------------
function settingRow_(label) {
  var sh = SpreadsheetApp.getActive().getSheetByName(TABS.SET);
  var labels = sh.getRange('A4:A30').getValues();
  for (var i = 0; i < labels.length; i++) if (labels[i][0] === label) return i + 4;
  return -1;
}
function settingValue_(label) {
  var r = settingRow_(label);
  return r > 0 ? SpreadsheetApp.getActive().getSheetByName(TABS.SET).getRange(r, 2).getValue() : '';
}
function setSetting_(label, value) {
  var r = settingRow_(label);
  if (r > 0) SpreadsheetApp.getActive().getSheetByName(TABS.SET).getRange(r, 2).setValue(value);
}
function asOf_() {
  var v = settingValue_('As-of date (in use)');
  return v instanceof Date ? v : new Date();
}

// ---------------------------------------------------------------------------------------------
// Duty Slots operations
// ---------------------------------------------------------------------------------------------
function slotsData_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(TABS.SLOTS);
  var last = sh.getLastRow();
  if (last < FIRST_ROW) return { sh: sh, rows: [] };
  return { sh: sh, rows: sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, S.SOURCE).getValues() };
}

function autoFillOpenSlots() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Auto-fill open G2 slots',
    'Fills every open, future G2 slot with the Soldier shown in "Suggested (next up)", one slot at a time ' +
    'in sheet order, so each pick updates the DA 6 before the next. Continue?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  sortSlots(true);
  var d = slotsData_();
  var today = stripTime_(asOf_());
  var filled = [], skipped = [];
  for (var i = 0; i < d.rows.length; i++) {
    var r = d.rows[i];
    if (!(r[S.DATE - 1] instanceof Date) || r[S.SOLDIER - 1] || r[S.STATUS - 1] === 'Cancelled') continue;
    if (stripTime_(r[S.DATE - 1]) < today) continue;
    if (ROSTERS.map(function (x) { return x.key; }).indexOf(r[S.ROSTER - 1]) < 0) continue;
    SpreadsheetApp.flush();
    var row = FIRST_ROW + i;
    var pick = String(d.sh.getRange(row, S.SUGGEST).getValue() || '');
    if (!pick || pick.charAt(0) === '⚠' || pick.charAt(0) === '(') {
      skipped.push(fmtDate_(r[S.DATE - 1]) + ' ' + r[S.DUTY - 1]);
      continue;
    }
    d.sh.getRange(row, S.SOLDIER).setValue(pick);
    d.sh.getRange(row, S.STATUS).setValue('Scheduled');
    var rem = d.sh.getRange(row, S.REMARKS);
    rem.setValue((rem.getValue() ? rem.getValue() + ' · ' : '') + 'auto-filled ' + fmtDate_(new Date()));
    filled.push(fmtDate_(r[S.DATE - 1]) + ' ' + r[S.DUTY - 1] + ' → ' + pick);
  }
  ui.alert('Auto-fill complete',
    (filled.length ? 'Filled:\n' + filled.join('\n') : 'Nothing to fill.') +
    (skipped.length ? '\n\nNo eligible Soldier available (fill manually):\n' + skipped.join('\n') : ''),
    ui.ButtonSet.OK);
}

function sortSlots(silent) {
  var d = slotsData_();
  if (!d.rows.length) return;
  var n = d.rows.filter(function (r) { return r[0] !== '' && r[0] !== null; }).length;
  // only input columns are sorted; computed columns are spilled formulas
  var inputs = [[1, 2], [4, 7], [14, 15]];
  var data = d.sh.getRange(FIRST_ROW, 1, d.rows.length, S.SOURCE).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] !== null; });
  data.sort(function (a, b) { return new Date(a[0]) - new Date(b[0]) || String(a[3]).localeCompare(String(b[3])); });
  inputs.forEach(function (span) {
    var w = span[1] - span[0] + 1;
    var block = data.map(function (r) { return r.slice(span[0] - 1, span[1]); });
    d.sh.getRange(FIRST_ROW, span[0], d.rows.length, w).clearContent();
    d.sh.getRange(FIRST_ROW, span[0], n, w).setValues(block);
  });
  if (silent !== true) SpreadsheetApp.getActive().toast('Duty Slots sorted by date.', 'G2 PSG', 4);
}

function markPastCompleted(silent) {
  var d = slotsData_();
  var today = stripTime_(new Date());
  var count = 0;
  d.rows.forEach(function (r, i) {
    var end = r[S.END - 1] instanceof Date ? r[S.END - 1] : r[S.DATE - 1];
    if (end instanceof Date && stripTime_(end) < today && r[S.SOLDIER - 1] && r[S.STATUS - 1] === 'Scheduled') {
      d.sh.getRange(FIRST_ROW + i, S.STATUS).setValue('Completed');
      count++;
    }
  });
  if (silent !== true) SpreadsheetApp.getActive().toast(count + ' slot(s) marked Completed.', 'G2 PSG', 4);
}

// ---------------------------------------------------------------------------------------------
// BDE roster import (parses the OCT26-style layout: SD / CQ tables, flag teams, road guard)
// ---------------------------------------------------------------------------------------------
function importFromRosterPrompt() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Import G2 slots',
    'Paste the BDE roster into its own tab (e.g. NOV26), then type that tab name:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var tab = res.getResponseText().trim();
  var sh = SpreadsheetApp.getActive().getSheetByName(tab);
  if (!sh) { ui.alert('No tab named "' + tab + '".'); return; }
  var unit = String(settingValue_('Unit label on BDE roster') || 'G2');
  var grid = sh.getDataRange().getDisplayValues();
  var people = personnelNames_();
  var found = parseBdeRoster(grid, unit, new Date().getFullYear()).map(function (s) {
    var m = matchName(s.rawName, s.rank, people);
    s.soldier = m.name;
    s.note = m.note;
    return s;
  });
  var d = slotsData_();
  var existing = {};
  d.rows.forEach(function (r) {
    if (r[0] instanceof Date) existing[slotKey_(r[0], r[S.DUTY - 1], r[S.POS - 1], r[S.SOLDIER - 1])] = true;
  });
  var fresh = found.filter(function (s) { return !existing[slotKey_(s.date, s.duty, s.position, s.soldier)]; });
  if (!fresh.length) { ui.alert('No new G2 slots found on "' + tab + '" (' + found.length + ' already tracked).'); return; }
  var start = Math.max(FIRST_ROW, lastDataRow_(d.sh) + 1);
  var out = fresh.map(function (s) {
    return [s.date, s.end || '', '', s.duty, s.position || '', s.soldier || '', s.soldier ? 'Scheduled' : ''];
  });
  d.sh.getRange(start, 1, out.length, 2).setValues(out.map(function (r) { return [r[0], r[1]]; }));
  d.sh.getRange(start, 4, out.length, 4).setValues(out.map(function (r) { return [r[3], r[4], r[5], r[6]]; }));
  d.sh.getRange(start, S.REMARKS, out.length, 2).setValues(fresh.map(function (s) {
    return [s.note || '', 'Import: ' + tab];
  }));
  sortSlots(true);
  var open = fresh.filter(function (s) { return !s.soldier; }).length;
  ui.alert('Import complete', fresh.length + ' G2 slot(s) added (' + open + ' open). ' +
    'Open slots show a suggested Soldier; run G2 PSG > Auto-fill open G2 slots to assign them.', ui.ButtonSet.OK);
}

/**
 * Pure parser (no Apps Script calls) so it can be unit tested.
 * grid: 2-D array of display strings. Returns [{date, end, duty, position, rank, rawName}].
 */
function parseBdeRoster(grid, unit, year) {
  var U = String(unit).toUpperCase().trim();
  var out = [];
  var H = grid.length, W = H ? grid[0].length : 0;
  var cell = function (r, c) { return (r >= 0 && r < H && c >= 0 && c < W) ? String(grid[r][c] || '').trim() : ''; };
  var up = function (r, c) { return cell(r, c).toUpperCase(); };

  // 1) SD / CQ tables: header row has Date ... Unit, Section, Rank, <name header>
  for (var r = 0; r < H; r++) {
    for (var c = 0; c < W; c++) {
      if (up(r, c) !== 'DATE' || up(r, c + 1) !== 'UNIT' || up(r, c + 2) !== 'SECTION') continue;
      var units = [];
      for (var k = c + 1; k < W; k++) if (up(r, k) === 'UNIT' && up(r, k + 1) === 'SECTION') units.push(k);
      units.forEach(function (u) {
        var title = '';
        for (var rr = r - 1; rr >= Math.max(0, r - 2) && !title; rr--) {
          for (var cc = u; cc >= Math.max(0, u - 6) && !title; cc--) if (cell(rr, cc)) title = cell(rr, cc);
        }
        var nameHdr = up(r, u + 3);
        var duty = classifyDuty_(title + ' ' + nameHdr);
        if (!duty) return;
        var blanks = 0;
        for (var rr2 = r + 1; rr2 < H && blanks < 3; rr2++) {
          var dtxt = cell(rr2, c);
          if (!dtxt) { blanks++; continue; }
          if (dtxt.toUpperCase() === 'DATE') break;
          blanks = 0;
          var dt = parseLooseDate(dtxt, year);
          if (!dt) continue;
          var sec = up(rr2, u + 1), unitCell = up(rr2, u);
          if (sec !== U && unitCell !== U) continue;
          var rank = up(rr2, u + 2);
          var name = cell(rr2, u + 3);
          var d2 = duty;
          if (duty === 'SD NCO' && OFFICER_RANKS.indexOf(rank) >= 0) d2 = 'SD OIC (Officer)';
          out.push({ date: dt, end: null, duty: d2, position: '', rank: rank, rawName: name });
        }
      });
    }
  }

  // 2) Flag-detail teams: "1-8OCT-26 (TEAM 1)" label in a DATE column; NAME / COMPANY / SECTION to the right
  for (var r3 = 0; r3 < H; r3++) {
    for (var c3 = 0; c3 < W; c3++) {
      var label = cell(r3, c3);
      var tm = /\(\s*TEAM\s*(\d+)\s*\)/i.exec(label);
      if (!tm) continue;
      var range = parseDateRange(label, year);
      if (!range) continue;
      // find header row above (DATE, NAME, COMPANY, SECTION)
      var hdrRow = -1;
      for (var h = r3 - 1; h >= Math.max(0, r3 - 15); h--) {
        if (up(h, c3) === 'DATE' && up(h, c3 + 1) === 'NAME') { hdrRow = h; break; }
      }
      var secCol = hdrRow >= 0 && up(hdrRow, c3 + 3) === 'SECTION' ? c3 + 3 : c3 + 3;
      for (var r4 = r3; r4 < H; r4++) {
        if (r4 > r3 && (up(r4, c3) === 'DATE' || /\(\s*TEAM/i.test(cell(r4, c3)))) break;
        if (r4 > r3 + 15) break;
        if (up(r4, secCol) !== U) continue;
        var raw = cell(r4, c3 + 1);
        var pos = 'Team ' + tm[1];
        var placeholder = /RANK\s+LAST/i.test(raw);
        if (placeholder) pos += ' · ' + raw.replace(/:?\s*RANK\s+LAST.*$/i, '').trim();
        out.push({ date: range[0], end: range[1], duty: 'Flag Detail', position: pos,
                   rank: placeholder ? '' : rankOf_(raw), rawName: placeholder ? '' : raw });
      }
    }
  }

  // 3) Road guard: "ROAD GUARD" title; rows below: date (carried), name, section
  for (var r5 = 0; r5 < H; r5++) {
    for (var c5 = 0; c5 < W; c5++) {
      if (up(r5, c5).indexOf('ROAD GUARD') < 0) continue;
      var curDate = null, blank = 0;
      for (var r6 = r5 + 1; r6 < H && blank < 3; r6++) {
        var dcell = cell(r6, c5), ncell = cell(r6, c5 + 1), scell = up(r6, c5 + 2);
        if (!dcell && !ncell && !scell) { blank++; continue; }
        blank = 0;
        if (dcell) { var pd = parseLooseDate(dcell, year); if (pd) curDate = pd; }
        if (!curDate || scell !== U) continue;
        out.push({ date: curDate, end: null, duty: 'Road Guard', position: '', rank: rankOf_(ncell), rawName: ncell });
      }
    }
  }
  return out;
}

function classifyDuty_(text) {
  var t = String(text).toUpperCase();
  if (t.indexOf('RUNNER') >= 0) return t.indexOf('CQ') >= 0 ? 'CQ Runner' : 'SD Runner';
  if (t.indexOf('CQ') >= 0 || t.indexOf('CHARGE OF QUARTERS') >= 0) return 'CQ NCO';
  if (t.indexOf('STAFF DUTY') >= 0 || t.indexOf('SDNCO') >= 0 || /\bSD\b/.test(t)) return 'SD NCO';
  return '';
}

var MONTHS_ = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };

/** "1-Oct-26", "2-Oct", "10/2/2026", "02 OCT 26" -> Date (no time) */
function parseLooseDate(text, year) {
  var t = String(text).trim().toUpperCase();
  var m = /^(\d{1,2})[\s\-]*([A-Z]{3})[A-Z]*[\s\-]*(\d{2,4})?$/.exec(t);
  if (m && MONTHS_[m[2]] !== undefined) {
    var y = m[3] ? (m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3])) : year;
    return new Date(y, MONTHS_[m[2]], Number(m[1]));
  }
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(t);
  if (m) {
    var yy = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
    return new Date(yy, Number(m[1]) - 1, Number(m[2]));
  }
  return null;
}

/** "1-8OCT-26 (TEAM 1)", "9-16OCT26", "25-31OCT26" -> [start, end] */
function parseDateRange(text, year) {
  var t = String(text).toUpperCase().replace(/\(.*\)/, '').trim();
  var m = /^(\d{1,2})\s*-\s*(\d{1,2})\s*([A-Z]{3})[A-Z]*\s*-?\s*(\d{2,4})?/.exec(t);
  if (!m || MONTHS_[m[3]] === undefined) return null;
  var y = m[4] ? (m[4].length === 2 ? 2000 + Number(m[4]) : Number(m[4])) : year;
  return [new Date(y, MONTHS_[m[3]], Number(m[1])), new Date(y, MONTHS_[m[3]], Number(m[2]))];
}

function rankOf_(text) {
  var first = String(text).trim().split(/\s+/)[0].toUpperCase();
  return RANKS.indexOf(first) >= 0 ? first : '';
}

/** Match "SGT DOE, JOHN" / a near-miss spelling like "DOH, JOHN" to a Personnel name. */
function matchName(raw, rank, people) {
  var t = String(raw || '').toUpperCase().replace(/\s+/g, ' ').trim();
  if (!t) return { name: '', note: '' };
  var parts = t.split(' ');
  if (RANKS.indexOf(parts[0]) >= 0) t = parts.slice(1).join(' ');
  t = t.replace(/\s*,\s*/, ', ');
  var norm = function (s) { return String(s).toUpperCase().replace(/[^A-Z]/g, ''); };
  for (var i = 0; i < people.length; i++) if (norm(people[i]) === norm(t)) return { name: people[i], note: '' };
  var last = norm(t.split(',')[0]);
  var first = norm((t.split(',')[1] || ''));
  var byLast = people.filter(function (p) { return norm(String(p).split(',')[0]) === last; });
  if (byLast.length === 1) return { name: byLast[0], note: '' };
  var best = null, bestD = 99;
  people.forEach(function (p) {
    var pl = norm(String(p).split(',')[0]), pf = norm(String(p).split(',')[1] || '');
    var dist = levenshtein_(pl, last) + (first && pf && pf.indexOf(first.slice(0, 3)) === 0 ? 0 : 2);
    if (dist < bestD) { bestD = dist; best = p; }
  });
  if (best && bestD <= 2) return { name: best, note: 'BDE roster: "' + raw + '" matched to ' + best + ' - verify' };
  return { name: '', note: 'BDE roster lists "' + raw + '" (not on Personnel) - add or assign' };
}

function levenshtein_(a, b) {
  var m = a.length, n = b.length, d = [];
  for (var i = 0; i <= m; i++) { d[i] = [i]; }
  for (var j = 0; j <= n; j++) { d[0][j] = j; }
  for (i = 1; i <= m; i++) for (j = 1; j <= n; j++) {
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  }
  return d[m][n];
}

function personnelNames_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(TABS.PERS);
  return sh.getRange('B4:B' + Math.max(4, sh.getLastRow())).getValues()
    .map(function (r) { return r[0]; }).filter(function (n) { return n && n !== 'VACANT'; });
}

function slotKey_(date, duty, pos, soldier) {
  return [fmtDate_(date), duty, pos || '', soldier || ''].join('|');
}

function lastDataRow_(sh) {
  var vals = sh.getRange(FIRST_ROW, 1, Math.max(1, sh.getMaxRows() - FIRST_ROW + 1), 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) if (vals[i][0] !== '' && vals[i][0] !== null) return FIRST_ROW + i;
  return FIRST_ROW - 1;
}

// ---------------------------------------------------------------------------------------------
// Shared view (no PII) - a separate spreadsheet you share with section sergeants
// ---------------------------------------------------------------------------------------------
function publishSharedView(silent) {
  var ss = SpreadsheetApp.getActive();
  var id = settingValue_('Shared view file ID');
  var target;
  try { target = id ? SpreadsheetApp.openById(id) : null; } catch (e) { target = null; }
  if (!target) {
    target = SpreadsheetApp.create('G2 Duty Roster - Shared View');
    var parent = DriveApp.getFileById(ss.getId()).getParents();
    if (parent.hasNext()) DriveApp.getFileById(target.getId()).moveTo(parent.next());
    setSetting_('Shared view file ID', target.getId());
  }
  var today = stripTime_(asOf_());
  var horizon = new Date(today.getTime() + 45 * 864e5);
  var d = slotsData_();
  var rows = [['Date', 'Day', 'Duty', 'Position', 'Rank', 'Soldier', 'Section', 'Status']];
  var sec = sectionsByName_();
  d.rows.forEach(function (r) {
    var dt = r[S.DATE - 1];
    if (!(dt instanceof Date) || dt < today || dt > horizon || r[S.STATUS - 1] === 'Cancelled') return;
    var who = r[S.SOLDIER - 1];
    rows.push([dt, Utilities.formatDate(dt, ss.getSpreadsheetTimeZone(), 'EEE'), r[S.DUTY - 1], r[S.POS - 1],
               r[S.RANK - 1] || '', who || 'OPEN - G2 to fill', who ? (sec[who] || '') : '', r[S.STATUS - 1] || 'Open']);
  });
  var sh = target.getSheets()[0].setName('Upcoming G2 Duties');
  sh.clear();
  sh.getRange(1, 1, 1, 8).merge().setValue('G2 duty roster - next 45 days (published ' +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd MMM yy HHmm') + ')')
    .setFontWeight('bold').setFontSize(13).setBackground('#1F3A5F').setFontColor('#FFFFFF');
  sh.getRange(2, 1, rows.length, 8).setValues(rows);
  sh.getRange(2, 1, 1, 8).setFontWeight('bold').setBackground('#5B6B7F').setFontColor('#FFFFFF');
  sh.getRange(3, 1, Math.max(1, rows.length - 1), 1).setNumberFormat('dd mmm yy');
  sh.setFrozenRows(2);
  sh.autoResizeColumns(1, 8);
  // next-up board: rank + name only (no numbers, no reasons)
  var nu = target.getSheetByName('Next Up') || target.insertSheet('Next Up');
  nu.clear();
  var board = [['Roster', '1st', '2nd', '3rd', '4th', '5th']];
  var nuSrc = ss.getSheetByName(TABS.NEXT);
  ROSTERS.forEach(function (r, idx) {
    var block = nuSrc.getRange(6, 1 + idx * 6, 5, 3).getDisplayValues();
    board.push([r.key].concat(block.map(function (b) { return b[2] ? (b[1] + ' ' + b[2]) : ''; })));
  });
  nu.getRange(1, 1, board.length, 6).setValues(board);
  nu.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
  nu.autoResizeColumns(1, 6);
  if (silent !== true) {
    SpreadsheetApp.getUi().alert('Shared view published',
      'File: ' + target.getUrl() + '\n\nShare THAT file (view-only) with section sergeants. ' +
      'It contains no phone numbers, absence reasons or DA 6 numbers, and refreshes every morning.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function sectionsByName_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(TABS.PERS);
  var vals = sh.getRange('A4:B' + Math.max(4, sh.getLastRow())).getValues();
  var m = {};
  vals.forEach(function (v) { if (v[1]) m[v[1]] = v[0]; });
  return m;
}

// ---------------------------------------------------------------------------------------------
// Exemption / availability request form -> Requests tab -> approve -> Availability
// ---------------------------------------------------------------------------------------------
function createRequestForm() {
  var ss = SpreadsheetApp.getActive();
  var existing = settingValue_('Exemption request form URL');
  if (existing) {
    SpreadsheetApp.getUi().alert('A request form already exists:\n' + existing);
    return;
  }
  var form = FormApp.create('G2 Duty Availability / Exemption Request');
  form.setDescription('Section leaders: submit Soldier leave, pass, appointments, school or exemption requests ' +
    'that affect G2 duty rosters. The G2 PSG reviews every request.');
  form.addListItem().setTitle('Soldier').setChoiceValues(personnelNames_()).setRequired(true);
  var types = ss.getSheetByName(TABS.SET).getRange('T4:T40').getValues()
    .map(function (r) { return r[0]; }).filter(String);
  form.addListItem().setTitle('Type').setChoiceValues(types).setRequired(true);
  form.addDateItem().setTitle('Start date').setRequired(true);
  form.addDateItem().setTitle('End date (leave blank if open-ended)');
  form.addParagraphTextItem().setTitle('Reason / details').setRequired(true);
  form.addTextItem().setTitle('Submitted by (rank / name / section)').setRequired(true);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  setSetting_('Exemption request form URL', form.getPublishedUrl());
  SpreadsheetApp.flush();
  ss.getSheets().forEach(function (s) { try { if (s.getFormUrl()) s.hideSheet(); } catch (err) {} });
  var req = ss.getSheetByName(TABS.REQ) || ss.insertSheet(TABS.REQ);
  req.clear();
  req.getRange(1, 1, 1, 10).setValues([['Received', 'Soldier', 'Type', 'Start', 'End', 'Reason',
    'Submitted by', 'Approve', 'Status', 'Form row']]).setFontWeight('bold').setBackground('#1F3A5F')
    .setFontColor('#FFFFFF');
  req.getRange('H2:H500').insertCheckboxes();
  req.setFrozenRows(1);
  req.setTabColor('#EDA100');
  ScriptApp.newTrigger('onRequestSubmit').forSpreadsheet(ss).onFormSubmit().create();
  SpreadsheetApp.getUi().alert('Request form created',
    'Share this link with section leaders:\n' + form.getPublishedUrl() + '\n\n' +
    'Submissions land on the "Requests" tab. Tick "Approve" to copy one to Availability.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function onRequestSubmit(e) {
  var ss = SpreadsheetApp.getActive();
  var req = ss.getSheetByName(TABS.REQ);
  if (!req || !e || !e.namedValues) return;
  var v = function (k) { return (e.namedValues[k] || [''])[0]; };
  req.appendRow([new Date(), v('Soldier'), v('Type'), v('Start date'), v('End date (leave blank if open-ended)'),
                 v('Reason / details'), v('Submitted by (rank / name / section)'), false, 'Pending', e.range.getRow()]);
  var row = req.getLastRow();
  req.getRange(row, 8).insertCheckboxes();
  req.getRange(row, 4, 1, 2).setNumberFormat('dd mmm yy');
}

function approveRequestRow_(sh, row, hdr) {
  var vals = sh.getRange(row, 1, 1, hdr.length).getValues()[0];
  var g = function (h) { return vals[hdr.indexOf(h)]; };
  if (String(g('Status')).indexOf('Approved') === 0) return;
  var av = SpreadsheetApp.getActive().getSheetByName(TABS.AVAIL);
  var target = lastDataRow_(av) + 1;
  av.getRange(target, 1, 1, 6).setValues([[g('Soldier'), g('Type'), toDate_(g('Start')), toDate_(g('End')) || '',
    g('Reason'), 'Request form: ' + g('Submitted by')]]);
  sh.getRange(row, hdr.indexOf('Status') + 1).setValue('Approved → Availability row ' + target);
}

// ---------------------------------------------------------------------------------------------
// Daily brief e-mail
// ---------------------------------------------------------------------------------------------
function emailBrief(silent) {
  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone();
  var bluf = ss.getSheetByName(TABS.DASH).getRange('I5').getDisplayValue();
  var today = stripTime_(asOf_());
  var wk = new Date(today.getTime() + 7 * 864e5);
  var d = slotsData_();
  var upcoming = [], conflicts = [];
  d.rows.forEach(function (r) {
    var dt = r[S.DATE - 1];
    if (!(dt instanceof Date) || r[S.STATUS - 1] === 'Cancelled') return;
    var line = Utilities.formatDate(dt, tz, 'EEE dd MMM') + ' · ' + r[S.DUTY - 1] + ' · ' +
      (r[S.SOLDIER - 1] ? (r[S.RANK - 1] + ' ' + r[S.SOLDIER - 1]) : '<b>OPEN</b>');
    if (dt >= today && dt <= wk) upcoming.push(line);
    if (dt >= today && String(r[S.CHECK - 1]).charAt(0) === '⚠' && r[S.SOLDIER - 1]) {
      conflicts.push(line + ' - ' + r[S.CHECK - 1]);
    }
  });
  var html = '<div style="font-family:Arial,sans-serif;font-size:13px">' +
    '<h3 style="margin:0 0 8px;color:#1F3A5F">G2 PSG brief · ' + Utilities.formatDate(today, tz, 'dd MMM yy') + '</h3>' +
    '<p><b>BLUF:</b> ' + bluf + '</p>' +
    '<p><b>Next 7 days</b><br>' + (upcoming.length ? upcoming.join('<br>') : 'No G2 duties.') + '</p>' +
    '<p><b>Conflicts</b><br>' + (conflicts.length ? conflicts.join('<br>') : 'None.') + '</p>' +
    '<p><a href="' + ss.getUrl() + '">Open the tracker</a>' +
    (settingValue_('One-slider file ID') ? ' · <a href="https://docs.google.com/presentation/d/' +
      settingValue_('One-slider file ID') + '/edit">Open the one-slider</a>' : '') + '</p></div>';
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(),
                      subject: 'G2 PSG brief · ' + Utilities.formatDate(today, tz, 'dd MMM yy'), htmlBody: html });
  if (silent !== true) ss.toast('Brief e-mailed to ' + Session.getEffectiveUser().getEmail(), 'G2 PSG', 5);
}

// ---------------------------------------------------------------------------------------------
// Add a tracker table (registered on Settings so it shows on the Dashboard)
// ---------------------------------------------------------------------------------------------
function addTrackerTablePrompt() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('Add tracker table', 'Table name (e.g. "Awards", "Counselings", "Weapons Quals"):',
    ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK || !a.getResponseText().trim()) return;
  var b = ui.prompt('Columns', 'Comma-separated column names. The first column should identify the row ' +
    '(e.g. "Soldier, Award, Submitted, Status, Due, Notes"):', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  var c = ui.prompt('Purpose', 'One line describing the table (shown on the Dashboard):', ui.ButtonSet.OK_CANCEL);
  addTrackerTable(a.getResponseText().trim(), b.getResponseText().split(',').map(function (s) { return s.trim(); })
    .filter(String), c.getSelectedButton() === ui.Button.OK ? c.getResponseText().trim() : '');
}

function addTrackerTable(name, columns, purpose) {
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(name)) throw new Error('A tab named "' + name + '" already exists.');
  if (!columns.length) columns = ['Item', 'Soldier', 'Status', 'Due', 'Notes'];
  var sh = ss.insertSheet(name, ss.getSheetByName(TABS.SET).getIndex() - 1);
  sh.setTabColor('#6250D6');
  sh.getRange(1, 1, 1, columns.length).merge().setValue(name.toUpperCase())
    .setFontSize(14).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F3A5F');
  sh.getRange(2, 1, 1, columns.length).merge().setValue(purpose || '')
    .setFontStyle('italic').setFontColor('#52514E');
  sh.getRange(3, 1, 1, columns.length).setValues([columns]).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#1F3A5F').setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(3);
  sh.getRange(3, 1, 998, columns.length).createFilter();
  sh.getRange(4, 1, 997, columns.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  columns.forEach(function (col, i) {
    var r = sh.getRange(4, i + 1, 997, 1);
    if (/soldier|name|holder/i.test(col)) {
      r.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(ss.getSheetByName(TABS.PERS).getRange('B4:B400'), true).setAllowInvalid(true).build());
    } else if (/date|due|suspense|submitted|received/i.test(col)) {
      r.setNumberFormat('dd mmm yy');
    } else if (/status/i.test(col)) {
      r.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['Open', 'In Progress', 'Complete', 'Closed'], true).setAllowInvalid(true).build());
    }
    sh.setColumnWidth(i + 1, Math.max(110, Math.min(260, col.length * 11)));
  });
  // register on Settings (AE:AG); the row count formula in AH is already there
  var set = ss.getSheetByName(TABS.SET);
  var reg = set.getRange('AE4:AE30').getValues();
  for (var k = 0; k < reg.length; k++) {
    if (!reg[k][0]) { set.getRange(4 + k, 31, 1, 3).setValues([[name, name, purpose || '']]); break; }
  }
  ss.setActiveSheet(sh);
  ss.toast('"' + name + '" added and registered on the Dashboard.', 'G2 PSG', 5);
  return sh;
}

// ---------------------------------------------------------------------------------------------
// Close & restart the DA 6s (AR 220-45 para 2a: "To" date = last detail; new roster brings numbers forward)
// ---------------------------------------------------------------------------------------------
function closeRosters() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Close & restart DA 6s',
    'This archives a values-only copy of every DA 6, copies each Soldier\'s current number into the ' +
    'Personnel "BF #" columns, and sets the roster FROM date to the day after the last posting. ' +
    'Do this when a DA 6 is nearly full (71 columns) or at a natural break. Continue?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone();
  var stamp = Utilities.formatDate(new Date(), tz, 'yyMMdd');
  var pers = ss.getSheetByName(TABS.PERS);
  var names = pers.getRange('B4:B400').getValues().map(function (r) { return r[0]; });
  var lastPosting = null;
  ROSTERS.forEach(function (r) {
    var sh = ss.getSheetByName(r.tab);
    var copy = sh.copyTo(ss).setName('ARCHIVE ' + r.key + ' ' + stamp);
    var rng = copy.getDataRange();
    rng.setValues(rng.getValues());
    copy.hideSheet();
    var block = sh.getRange('B8:H127').getValues();
    block.forEach(function (row) {
      var i = names.indexOf(row[0]);
      if (row[0] && i >= 0 && typeof row[6] === 'number') pers.getRange(4 + i, r.bfCol).setValue(row[6]);
    });
    sh.getRange('I4:CA4').getValues()[0].forEach(function (v) {
      if (typeof v === 'number' && v > 0) v = new Date(Math.round((v - 25569) * 864e5) + new Date().getTimezoneOffset() * 6e4);
      if (v instanceof Date && (!lastPosting || v > lastPosting)) lastPosting = v;
    });
  });
  var from = lastPosting ? new Date(lastPosting.getTime() + 864e5) : stripTime_(new Date());
  setSetting_('Roster FROM date', from);
  setSetting_('Close-out (archive) log', 'Closed ' + Utilities.formatDate(new Date(), tz, 'dd MMM yy') +
    '; new FROM ' + Utilities.formatDate(from, tz, 'dd MMM yy') + '; archives "ARCHIVE * ' + stamp + '" (hidden)');
  ui.alert('DA 6s closed', 'Archived (hidden tabs "ARCHIVE … ' + stamp + '"). New FROM date: ' +
    Utilities.formatDate(from, tz, 'dd MMM yy') + '. Brought-forward numbers are on Personnel.', ui.ButtonSet.OK);
}

// ---------------------------------------------------------------------------------------------
// Mind map (org chart) and chart rebuild
// ---------------------------------------------------------------------------------------------
function buildMindMap() {
  var ss = SpreadsheetApp.getActive();
  var dash = ss.getSheetByName(TABS.DASH);
  dash.getCharts().forEach(function (c) {
    if (c.getOptions().get('title') === 'G2 PSG mind map') dash.removeChart(c);
  });
  var src = ss.getSheetByName(TABS.DATA).getRange('AG5:AH25');
  var chart = dash.newChart().setChartType(Charts.ChartType.ORG).addRange(src)
    .setPosition(103, 2, 0, 0).setOption('title', 'G2 PSG mind map').setOption('size', 'medium')
    .setOption('allowHtml', true).setOption('width', 1250).setOption('height', 380).build();
  dash.insertChart(chart);
}

function rebuildCharts() {
  var ss = SpreadsheetApp.getActive();
  var dash = ss.getSheetByName(TABS.DASH);
  var dd = ss.getSheetByName(TABS.DATA);
  var keep = 'G2 PSG mind map';
  dash.getCharts().forEach(function (c) { if (c.getOptions().get('title') !== keep) dash.removeChart(c); });
  var palette = ['#2A78D6', '#EB6834', '#1BAF7A', '#EDA100', '#E87BA4', '#008300', '#6250D6', '#E34948'];
  var specs = [
    { type: Charts.ChartType.COLUMN, range: 'A4:G12', title: 'G2 duty details by section', row: 39, col: 2, stacked: true,
      colors: palette.slice(0, 5) },
    { type: Charts.ChartType.COLUMN, range: ['H4:H10', 'J4:K10'], title: 'This month: filled vs open G2 slots', row: 45,
      col: 9, stacked: true, colors: ['#2A78D6', '#C9CED6'] },
    { type: Charts.ChartType.LINE, range: 'O4:P34', title: 'Soldiers unavailable, next 30 days', row: 57, col: 2,
      colors: ['#2A78D6'], legend: 'none' },
    { type: Charts.ChartType.BAR, range: 'U4:V19', title: 'Duty load per Soldier (top 15): fairness check', row: 63,
      col: 9, colors: ['#2A78D6'], legend: 'none' },
    { type: Charts.ChartType.BAR, range: 'R4:S17', title: 'Availability records this month, by type', row: 75, col: 2,
      colors: ['#1BAF7A'], legend: 'none' },
    { type: Charts.ChartType.COLUMN, range: 'X4:Y9', title: 'Equipment by status', row: 84, col: 9,
      colors: ['#6250D6'], legend: 'none' },
    { type: Charts.ChartType.COLUMN, range: 'AA4:AB9', title: 'Taskings by status', row: 95, col: 2,
      colors: ['#EDA100'], legend: 'none' }
  ];
  specs.forEach(function (s) {
    var b = dash.newChart().setChartType(s.type);
    (Array.isArray(s.range) ? s.range : [s.range]).forEach(function (a1) { b.addRange(dd.getRange(a1)); });
    b.setNumHeaders(1).setPosition(s.row, s.col, 0, 0).setOption('title', s.title).setOption('colors', s.colors)
      .setOption('legend', { position: s.legend || 'bottom' }).setOption('width', 620).setOption('height', 320)
      .setOption('backgroundColor', '#FFFFFF');
    if (s.stacked) b.setOption('isStacked', true);
    dash.insertChart(b.build());
  });
}

// ---------------------------------------------------------------------------------------------
// One-slider brief for senior leaders (Google Slides, one 16:9 slide, rebuilt by the daily job)
// ---------------------------------------------------------------------------------------------
var SLIDE_CHART_TITLE = 'One-slider: filled vs open G2 slots';

function buildOneSlider(silent) {
  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone();
  var dash = ss.getSheetByName(TABS.DASH);
  var dd = ss.getSheetByName(TABS.DATA);
  var today = stripTime_(asOf_());

  // deck: reuse the one on Settings, else create it next to the tracker
  var deck = null, id = settingValue_('One-slider file ID');
  if (id) { try { deck = SlidesApp.openById(id); } catch (err) { deck = null; } }
  if (!deck) {
    deck = SlidesApp.create('G2 PSG One-Slider');
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) DriveApp.getFileById(deck.getId()).moveTo(parents.next());
    setSetting_('One-slider file ID', deck.getId());
  }
  var slides = deck.getSlides();
  var slide = slides.length ? slides[0] : deck.appendSlide();
  for (var i = 1; i < slides.length; i++) slides[i].remove();
  slide.getPageElements().forEach(function (el) { el.remove(); });
  slide.getBackground().setSolidFill('#FCFCFB');

  var W = deck.getPageWidth(), M = 18, inner = W - 2 * M;
  var marking = String(settingValue_('Brief marking (optional)') || '').trim();
  var psg = String(settingValue_('PSG') || '');

  // header
  slideBox_(slide, 0, 0, W, 50, { fill: '#1F3A5F' });
  slideBox_(slide, M, 8, inner * 0.6, 34, { text: 'G2 DUTY & PERSONNEL BRIEF', size: 18, bold: true, color: '#FFFFFF' });
  slideBox_(slide, M + inner * 0.5, 8, inner * 0.5, 34, {
    text: 'As of ' + Utilities.formatDate(today, tz, 'dd MMM yy') + (psg ? '\n' + psg + ' · G2 PSG' : ''),
    size: 9, color: '#FFFFFF', align: SlidesApp.ParagraphAlignment.END });
  if (marking) slideBox_(slide, 0, 0, W, 10, { text: marking, size: 7, bold: true, color: '#FFFFFF',
    align: SlidesApp.ParagraphAlignment.CENTER });

  // BLUF
  slideBox_(slide, M, 58, inner, 16, { text: 'BLUF', size: 9, bold: true, color: '#1F3A5F' });
  slideBox_(slide, M, 74, inner, 50, { text: dash.getRange('I5').getDisplayValue(), size: 10.5, color: '#0B0B0B',
    fill: '#FFFFFF', line: '#D5DAE1', valign: SlidesApp.ContentAlignment.TOP });

  // KPI tiles: [dashboard value cell, label, status rule on Dash Data]
  var kpi = function (a1) { return dd.getRange(a1).getValue(); };
  var tiles = [
    ['B5', 'ASSIGNED', null],
    ['D5', 'PRESENT FOR DUTY', null],
    ['F5', 'G2 SLOTS FILLED (MONTH)', kpi('AE12') < 1 ? '#C77700' : null],
    ['B9', 'OPEN SLOTS (LOOK-AHEAD)', kpi('AE13') > 0 ? '#D03B3B' : null],
    ['D9', 'CONFLICTS', kpi('AE14') > 0 ? '#C77700' : null],
    ['B13', 'OPEN TASKINGS', kpi('AE16') > 0 ? '#D03B3B' : null]
  ];
  var gap = 8, tw = (inner - gap * (tiles.length - 1)) / tiles.length, ty = 132;
  tiles.forEach(function (t, k) {
    var x = M + k * (tw + gap);
    slideBox_(slide, x, ty, tw, 64, { fill: '#FFFFFF', line: '#D5DAE1' });
    slideBox_(slide, x, ty + 3, tw, 14, { text: t[1], size: 7, bold: true, color: '#52514E',
      align: SlidesApp.ParagraphAlignment.CENTER });
    slideBox_(slide, x, ty + 16, tw, 28, { text: dash.getRange(t[0]).getDisplayValue(), size: 20, bold: true,
      color: t[2] || '#1F3A5F', align: SlidesApp.ParagraphAlignment.CENTER });
    var note = dash.getRange(t[0]).offset(1, 0).getDisplayValue();
    slideBox_(slide, x + 3, ty + 44, tw - 6, 18, { text: note, size: 6.5, color: '#52514E',
      align: SlidesApp.ParagraphAlignment.CENTER });
  });

  // next 7 days of G2 duties
  var by = 206, half = (inner - 12) / 2;
  slideBox_(slide, M, by, half, 16, { text: 'G2 DUTIES · NEXT 7 DAYS', size: 9, bold: true, color: '#1F3A5F' });
  var wk = new Date(today.getTime() + 7 * 864e5);
  var rows = slotsData_().rows.filter(function (r) {
    var dt = r[S.DATE - 1];
    return dt instanceof Date && dt >= today && dt <= wk && r[S.STATUS - 1] !== 'Cancelled';
  }).sort(function (a, b) { return a[S.DATE - 1] - b[S.DATE - 1]; });
  var MAXL = 10, lines = rows.slice(0, MAXL).map(function (r) {
    var who = r[S.SOLDIER - 1] ? (r[S.RANK - 1] + ' ' + r[S.SOLDIER - 1] + (r[S.SECTION - 1] ? ' (' + r[S.SECTION - 1] + ')' : ''))
                               : 'OPEN - needs a Soldier';
    return Utilities.formatDate(r[S.DATE - 1], tz, 'EEE dd MMM') + '   ' + r[S.DUTY - 1] + '   ' + who;
  });
  if (rows.length > MAXL) lines.push('+' + (rows.length - MAXL) + ' more on the tracker');
  var list = slideBox_(slide, M, by + 18, half, 162, { text: lines.length ? lines.join('\n') : 'No G2 duties in the next 7 days.',
    size: 8.5, color: '#0B0B0B', fill: '#FFFFFF', line: '#D5DAE1', valign: SlidesApp.ContentAlignment.TOP });
  var txt = list.getText(), pos = 0;
  lines.forEach(function (l) {
    var at = l.indexOf('OPEN - needs a Soldier');
    if (at >= 0) txt.getRange(pos + at, pos + l.length).getTextStyle().setBold(true).setForegroundColor('#D03B3B');
    pos += l.length + 1;
  });

  // linked (live) chart: rebuilt on Dash Data so the slide never depends on Dashboard chart titles
  dd.getCharts().forEach(function (c) { if (c.getOptions().get('title') === SLIDE_CHART_TITLE) dd.removeChart(c); });
  var chart = dd.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(dd.getRange('H4:H10')).addRange(dd.getRange('J4:K10')).setNumHeaders(1)
    .setPosition(62, 1, 0, 0).setOption('title', SLIDE_CHART_TITLE).setOption('isStacked', true)
    .setOption('colors', ['#2A78D6', '#C9CED6']).setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', '#FFFFFF').setOption('width', 520).setOption('height', 300).build();
  dd.insertChart(chart);
  chart = dd.getCharts().filter(function (c) { return c.getOptions().get('title') === SLIDE_CHART_TITLE; })[0];
  var cx = M + half + 12;
  slideBox_(slide, cx, by, half, 16, { text: 'THIS MONTH · FILLED VS OPEN BY DA 6', size: 9, bold: true, color: '#1F3A5F' });
  slide.insertSheetsChart(chart, cx, by + 18, half, 162);

  // footer
  slideBox_(slide, M, 388, inner, 14, {
    text: 'Source: G2 PSG Tracker (live) · refreshed daily 0500 · ' + Utilities.formatDate(new Date(), tz, 'dd MMM yy HHmm') +
      (marking ? '   |   ' + marking : ''), size: 6.5, color: '#5B6B7F' });
  deck.saveAndClose();
  if (silent !== true) {
    SpreadsheetApp.getUi().alert('One-slider ready',
      'https://docs.google.com/presentation/d/' + deck.getId() + '/edit\n\nIt sits next to the tracker and the daily job refreshes it.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
  return deck.getId();
}

/** Text box / rectangle helper for the one-slider. */
function slideBox_(slide, x, y, w, h, o) {
  var sh = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  if (o.fill) sh.getFill().setSolidFill(o.fill); else sh.getFill().setTransparent();
  if (o.line) sh.getBorder().setWeight(0.75).getLineFill().setSolidFill(o.line); else sh.getBorder().setTransparent();
  sh.setContentAlignment(o.valign || SlidesApp.ContentAlignment.MIDDLE);
  if (o.text) {
    var tr = sh.getText();
    tr.setText(String(o.text));
    tr.getTextStyle().setFontFamily('Arial').setFontSize(o.size || 10).setBold(!!o.bold)
      .setForegroundColor(o.color || '#0B0B0B');
    tr.getParagraphStyle().setParagraphAlignment(o.align || SlidesApp.ParagraphAlignment.START);
  }
  return sh;
}

// ---------------------------------------------------------------------------------------------
// small utils
// ---------------------------------------------------------------------------------------------
function stripTime_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function toDate_(v) { return v instanceof Date ? v : (v ? new Date(v) : null); }
function fmtDate_(d) {
  if (!(d instanceof Date)) return String(d || '');
  return Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'dd MMM yy');
}

// export pure functions for local unit tests (ignored by Apps Script)
if (typeof module !== 'undefined') {
  module.exports = { parseBdeRoster: parseBdeRoster, parseLooseDate: parseLooseDate, parseDateRange: parseDateRange,
                     matchName: matchName, classifyDuty_: classifyDuty_ };
}
