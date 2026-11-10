// Local check of the BDE-roster importer in Code.gs against a mock of the OCT26 tab layout.
// Run: node psg-tracker/tests/parser.check.cjs
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
const ctx = { module: { exports: {} }, console };
vm.createContext(ctx);
vm.runInContext(code, ctx);
const { parseBdeRoster, matchName, parseLooseDate, parseDateRange } = ctx.module.exports;

// --- mock grid (display values) in the OCT26 layout: A..X, row 1 = index 0 -----------------
const W = 24, H = 70;
const g = Array.from({ length: H }, () => Array(W).fill(''));
const C = (a1) => { const m = /^([A-Z]+)(\d+)$/.exec(a1); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [Number(m[2]) - 1, c - 1]; };
const put = (a1, v) => { const [r, c] = C(a1); g[r][c] = v; };
const row = (r, startCol, vals) => vals.forEach((v, i) => { const [, c] = C(startCol + '1'); g[r - 1][c + i] = v; });

put('A1', 'SD Roster');
put('A2', 'Staff Duty OIC/NCOIC'); put('H2', 'Staff Duty Runner');
row(2, 'O', ['DATE', 'NAME', 'COMPANY', 'SECTION', 'NUMBER', 'DATE', 'NAME', 'COMPANY', 'SECTION', 'NUMBER']);
row(3, 'A', ['Date', 'Unit', 'Section', 'Rank', 'SDNCO BLDG. 2025', 'PHONE NUMBER', '', 'Unit', 'Section', 'Rank', 'SD Runner BLDG. 2025', 'PHONE NUMBER']);
const sd = [
  ['1-Oct-26', 'HSC', '', '', '', '', '', 'HSC', '', '', ''],
  ['2-Oct-26', 'SIS', 'G2', 'CPT', 'YATES, SAM', '', '', 'SIS', '1ST', 'PFC', 'ABBOTT, CHRIS'],
  ['13-Oct-26', 'SIS', 'SURG', 'SSG', 'TELLER, WILL', '', '', 'SIS', 'G2', 'PV2', 'SHORT, REED'],
  ['14-Oct-26', 'SIS', 'G2', 'SFC', 'GARZA, CHRIS', '', '', 'HSC', '', '', ''],
  ['20-Oct-26', 'BAND', '', '', '', '', '', 'SIS', 'G2', 'SPC', ''],
  ['24-Oct-26', 'SIS', 'G2', 'CW2', 'PARKER, DAN', '', '', 'SIS', '1ST', 'SPC', 'NASH, JOHN'],
];
sd.forEach((v, i) => row(4 + i, 'A', v));
// flag teams
put('O3', '1-8OCT-26 (TEAM 1)');
row(3, 'P', ['SFC Doe, Pat', 'SIS', 'G6', '(555) 010-0100']);
row(4, 'P', ['SGT STONE, TAYLOR', 'SIS', 'G2', '']);
row(5, 'P', ['SGT TATE, CORY', 'SIS', '1ST PLT', '']);
row(6, 'P', ['E5-E6: Rank Last, First', 'SIS', 'G2', '']);
put('T3', '9-16OCT26 (TEAM 2)'); row(3, 'U', ['NCOIC: E6-E7: Rank Last, First', 'HSC', '', '']);
row(14, 'O', ['DATE', 'NAME', 'COMPANY', 'SECTION', 'NUMBER']);
put('O15', '17-24OCT26 (TEAM 3)'); row(15, 'P', ['SSG LANGLEY, DAVE', 'SIS', 'G2', '555-010-0101']);
row(16, 'P', ['SPC SATO, BEN', 'SIS', 'G2', '555-010-0102']);
// CQ table further down
put('A36', 'CQ Roster'); put('A37', 'CQ NCOIC'); put('H37', 'CQ Runner');
row(38, 'A', ['Date', 'Unit', 'Section', 'Rank', 'CQNCO BLDG. 2025', 'PHONE NUMBER', '', 'Unit', 'Section', 'Rank', 'CQ BLDG 2020 RUNNER', 'PHONE NUMBER']);
row(39, 'A', ['2-Oct-26', 'SIS', 'G2', 'SGT', 'HARTMAN, BEN', '', '', 'SIS', 'G6', 'SPC', 'RHODES, CONNOR']);
row(40, 'A', ['12-Oct-26', 'SIS', 'HQ PLT', 'SGT', 'ELLISON, LEE', '', '', 'SIS', 'G2', 'SPC', 'QUINCY, CHRIS']);
row(41, 'A', ['30-Oct-26', 'SIS', 'G2', 'SGT', 'MORLAES, MAX', '', '', 'BAND', '', '', '']);
// road guard
put('O27', 'ROAD GUARD DETAIL');
row(28, 'O', ['2-Oct', '', '1ST PLT']); row(29, 'O', ['', 'SPC VEGARA', '2ND PLT']); row(30, 'O', ['', 'PFC BAKER, TOM', 'G2']);
row(31, 'O', ['6-Oct', 'SPC DRAKE, SHANE', 'G6']);
row(34, 'O', ['30-Oct', '', '1ST PLT']); row(35, 'O', ['', 'SPC Calixto', '2ND PLT']); row(36, 'O', ['', '', 'G2']);

const slots = parseBdeRoster(g, 'G2', 2026);
const key = (s) => `${s.date.toISOString().slice(0, 10)}|${s.duty}|${s.position}|${s.rawName}`;
const got = slots.map(key).sort();
const expected = [
  '2026-10-02|SD OIC (Officer)||YATES, SAM',
  '2026-10-13|SD Runner||SHORT, REED',
  '2026-10-14|SD NCO||GARZA, CHRIS',
  '2026-10-20|SD Runner||',
  '2026-10-24|SD OIC (Officer)||PARKER, DAN',
  '2026-10-01|Flag Detail|Team 1|SGT STONE, TAYLOR',
  '2026-10-01|Flag Detail|Team 1 · E5-E6|',
  '2026-10-17|Flag Detail|Team 3|SSG LANGLEY, DAVE',
  '2026-10-17|Flag Detail|Team 3|SPC SATO, BEN',
  '2026-10-02|CQ NCO||HARTMAN, BEN',
  '2026-10-12|CQ Runner||QUINCY, CHRIS',
  '2026-10-30|CQ NCO||MORLAES, MAX',
  '2026-10-02|Road Guard||PFC BAKER, TOM',
  '2026-10-30|Road Guard||',
].sort();
assert.strictEqual(JSON.stringify(got), JSON.stringify(expected));
const flagEnd = slots.find((s) => s.position === 'Team 3' && s.rawName.includes('LANGLEY')).end;
assert.strictEqual(flagEnd.getDate(), 24);

const people = ['HARTMANN, BEN', 'MORALES, MAX', 'STONE, TAYLOR', 'SHORT, REED', 'BAKER, TOM', 'PARKER, DAN', 'PARKER, SAM'];
assert.strictEqual(matchName('SGT STONE, TAYLOR', 'SGT', people).name, 'STONE, TAYLOR');
assert.strictEqual(matchName('HARTMAN, BEN', 'SGT', people).name, 'HARTMANN, BEN');
assert.strictEqual(matchName('MORLAES, MAX', 'SGT', people).name, 'MORALES, MAX');
assert.strictEqual(matchName('PARKER, DAN', 'CW2', people).name, 'PARKER, DAN');
assert.strictEqual(matchName('YATES, SAM', 'CPT', people).name, '');
assert.strictEqual(parseLooseDate('2-Oct', 2026).getMonth(), 9);
assert.strictEqual(parseDateRange('9-16OCT26 (TEAM 2)', 2026)[1].getDate(), 16);
console.log('parser check passed:', slots.length, 'G2 slots');
console.log(got.join('\n'));
