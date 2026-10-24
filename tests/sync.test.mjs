import { test } from 'node:test';
import assert from 'node:assert/strict';

import { createNameMatcher } from '../js/names.js';
import { parseCells } from '../js/schedule-parser.js';
import { parseRosterRows } from '../js/roster.js';
import { mergeSchedule } from '../js/scheduling.js';
import { stampChanges, mergeStates, describeChanges, packState, unpackState, erasedSince } from '../js/sync.js';

const rows = [
  ['שם', 'טלפון', 'הדרכה', 'הדרכת בטיחות', 'הערות'],
  ['דנה לוי', '1', 'V', 'V', 'קוסמת'],
  ['עומר בן חיים', '2', 'V', 'V', 'עושה הרצאות'],
  ['רון אלמוג', '4', 'V', 'V', ''],
  ['תמר גל', '5', 'V', 'V', ''],
];
const instructors = parseRosterRows(rows, new Date('2026-10-24'));
const matcher = createNameMatcher(instructors.map(i => i.name));

const cells = [
  { date: '2026-11-02', lines: ['יום עיון בכימיה לתיכון א', '10:00-11:30 - מעבדה בכימיה. מדריכים: ?+?', '12:00-12:45 - מופע קסמים במדע. מדריך: ?'] },
  { date: '2026-11-03', lines: ['יום עיון בכימיה לתיכון ב', '10:30-12:00 - מעבדה בכימיה. מדריכים: ?+?'] },
];

const T0 = Date.parse('2026-10-24T08:00:00Z');
const at = minutes => T0 + minutes * 60000;
const clone = o => JSON.parse(JSON.stringify(o));

/** A device: its state, and commit/receive the way the page does them. */
function device(settings = { month: 'all' }) {
  const d = {
    state: {
      version: 1, instructors: clone(instructors), extraInstructors: [], instructorEdits: {}, events: [],
      settings, notes: { gaps: '', tasks: '' }, exportOptions: {},
      imports: { roster: { name: 'r.xlsx', at: new Date(at(0)).toISOString(), count: 4 }, schedule: [] }, lastImport: null,
    },
    snap: null,
    commit(t) { stampChanges(d.snap, d.state, t); d.snap = clone(d.state); },
    receive(other, t) { d.state = mergeStates(d.state, unpackState(clone(packState(other.state, 'x'))), t); d.snap = clone(d.state); },
    import(list, t) {
      const parsed = parseCells(list, matcher);
      d.state.events = mergeSchedule(d.state.events, parsed).events;
      const dates = parsed.map(e => e.date).sort();
      d.state.imports.schedule.push({ name: 'lz.pdf', at: new Date(t).toISOString(), events: parsed.length, from: dates[0], to: dates[dates.length - 1] });
      d.commit(t);
    },
    pos(title, si, pi, s = d.state) { return s.events.find(e => e.title.includes(title)).slots[si].groups[0].positions[pi]; },
    assign(title, si, pi, name, t) { Object.assign(d.pos(title, si, pi), { name, source: name ? 'user' : null }); d.commit(t); },
  };
  d.snap = clone(d.state);
  return d;
}

test('each device keeps its own assignments after a merge, both ways', () => {
  const a = device(), b = device();
  a.import(cells, at(1));
  b.import(cells, at(2));
  a.assign('תיכון א', 0, 0, 'רון אלמוג', at(3));
  b.assign('תיכון ב', 0, 1, 'תמר גל', at(4));
  a.receive(b, at(5));
  b.receive(a, at(6));
  for (const d of [a, b]) {
    assert.equal(d.pos('תיכון א', 0, 0).name, 'רון אלמוג');
    assert.equal(d.pos('תיכון ב', 0, 1).name, 'תמר גל');
  }
  // Receiving the same thing again changes nothing.
  const before = JSON.stringify(a.state.events);
  a.receive(b, at(7));
  assert.equal(JSON.stringify(a.state.events), before);
});

test('the same position changed on both devices: the later choice wins, also a clear', () => {
  const a = device(), b = device();
  a.import(cells, at(1));
  b.import(cells, at(1));
  a.assign('תיכון א', 0, 0, 'רון אלמוג', at(3));
  b.assign('תיכון א', 0, 0, 'דנה לוי', at(4));
  a.receive(b, at(5));
  assert.equal(a.pos('תיכון א', 0, 0).name, 'דנה לוי');
  // A clears it later; B gets the clear even though B still has a name there.
  a.assign('תיכון א', 0, 0, null, at(6));
  b.receive(a, at(7));
  assert.equal(b.pos('תיכון א', 0, 0).name, null);
});

test('a newer schedule on the other device gives the shape; older assignments ride on it', () => {
  const a = device(), b = device();
  a.import(cells, at(1));
  b.import(cells, at(1));
  a.assign('תיכון א', 0, 1, 'עומר בן חיים', at(2));
  // B imports an update: the title gained a detail and the lab moved by half an hour.
  const updated = [
    { date: '2026-11-02', lines: ['יום עיון בכימיה לתיכון א (רכזת: 050)', '10:30-12:00 - מעבדה בכימיה. מדריכים: ?+?', '12:00-12:45 - מופע קסמים במדע. מדריך: ?'] },
    cells[1],
  ];
  b.import(updated, at(10));
  a.receive(b, at(11));
  const ev = a.state.events.find(e => e.title.includes('תיכון א'));
  assert.match(ev.title, /רכזת/);
  assert.equal(ev.slots[0].start, '10:30');
  assert.equal(ev.slots[0].groups[0].positions[1].name, 'עומר בן חיים');
  assert.equal(a.state.events.length, 2);
});

test('a name the coordinator wrote in a newer schedule beats an older choice', () => {
  const a = device(), b = device();
  a.import(cells, at(1));
  b.import(cells, at(1));
  a.assign('תיכון ב', 0, 0, 'רון אלמוג', at(2));
  b.import([cells[0], { date: '2026-11-03', lines: ['יום עיון בכימיה לתיכון ב', '10:30-12:00 - מעבדה בכימיה. מדריכים: תמר גל+?'] }], at(10));
  a.receive(b, at(11));
  assert.equal(a.pos('תיכון ב', 0, 0).name, 'תמר גל');
});

test('manual events travel, and a deletion on one device holds on the other', () => {
  const a = device(), b = device();
  a.import(cells, at(1));
  b.import(cells, at(1));
  a.state.events.push({ id: 'mtest', date: '2026-11-10', color: null, kind: 'study', subject: null, title: 'יום עיון ידני', raw: '', manual: true, ignored: false, note: '', changes: [],
    slots: [{ id: 'mtest.m1', start: '10:00', end: '11:00', kind: 'lab', label: 'הדרכה', location: '', subject: null, manual: true, edited: true, groups: [{ label: '', activity: '', positions: [{ name: 'דנה לוי', source: 'user' }, { name: null, source: null }] }] }] });
  a.commit(at(2));
  b.receive(a, at(3));
  assert.equal(b.pos('ידני', 0, 0).name, 'דנה לוי');
  b.state.events = b.state.events.filter(e => e.id !== 'mtest');
  b.state.deleted = { events: { mtest: at(4) }, slots: {} };
  b.commit(at(4));
  a.receive(b, at(5));
  assert.equal(a.state.events.some(e => e.id === 'mtest'), false);
});

test('notes, instructor edits and the roster: the newer one wins; view settings stay per device', () => {
  const a = device({ month: '2026-11' }), b = device({ month: 'all' });
  a.state.instructorEdits['רון אלמוג'] = { active: false };
  a.commit(at(1));
  b.state.instructorEdits['רון אלמוג'] = { active: true, notes: 'חזר' };
  b.state.notes.tasks = 'לסגור את נובמבר';
  b.commit(at(2));
  b.state.imports.roster = { name: 'r2.xlsx', at: new Date(at(3)).toISOString(), count: 3 };
  b.state.instructors = b.state.instructors.slice(0, 3);
  b.commit(at(3));
  const before = clone(a.state);
  a.receive(b, at(4));
  assert.equal(a.state.instructorEdits['רון אלמוג'].notes, 'חזר');
  assert.equal(a.state.notes.tasks, 'לסגור את נובמבר');
  assert.equal(a.state.instructors.length, 3);
  assert.equal(a.state.settings.month, '2026-11');
  const d = describeChanges(before, a.state);
  assert.equal(d.roster, true);
  assert.equal(d.notes, true);
  assert.equal(d.people, 1);
});

test('events outside the newer file stay, and the summary counts what arrived', () => {
  const a = device(), b = device();
  a.import([{ date: '2026-10-28', lines: ['יום עיון בכימיה לתיכון ז', '10:00-11:00 - מעבדה בכימיה. מדריכים: ?+?'] }], at(1));
  a.assign('תיכון ז', 0, 0, 'רון אלמוג', at(2));
  b.import(cells, at(5));
  const before = clone(b.state);
  b.receive(a, at(6));
  assert.equal(b.pos('תיכון ז', 0, 0).name, 'רון אלמוג');
  const d = describeChanges(before, b.state);
  assert.equal(d.events, 1);
  assert.equal(d.assigned, 1);
  assert.equal(d.schedule, false);
});

test('an old backup file (bare state, no stamps) is still accepted', () => {
  const a = device();
  a.import(cells, at(1));
  const legacy = clone(a.state);
  delete legacy.settings;
  assert.ok(unpackState(legacy));
  assert.equal(unpackState({ hello: 1 }), null);
});

test('"delete everything" wins over older data in both directions, and a new start merges normally', () => {
  const a = device(), b = device();
  a.import(cells, at(1));
  b.import(cells, at(1));
  a.assign('תיכון א', 0, 0, 'רון אלמוג', at(2));
  b.assign('תיכון ב', 0, 0, 'תמר גל', at(3));
  // A erases everything (a fresh state with a new generation).
  a.state = { ...device().state, events: [], instructors: [], imports: { roster: null, schedule: [] }, resetAt: at(5) };
  a.snap = clone(a.state);
  assert.equal(erasedSince(b.state, a.state), true);
  // B's older work does not come back into A...
  a.receive(b, at(6));
  assert.equal(a.state.events.length, 0);
  // ...and B, receiving A, is erased too (its view settings stay).
  b.state.settings = { month: '2026-11' };
  b.receive(a, at(7));
  assert.equal(b.state.events.length, 0);
  assert.equal(b.state.settings.month, '2026-11');
  // After the erase both start over: same generation, ordinary merge.
  a.import(cells, at(10));
  a.assign('תיכון א', 0, 1, 'דנה לוי', at(11));
  b.import(cells, at(12));
  b.assign('תיכון ב', 0, 1, 'עומר בן חיים', at(13));
  a.receive(b, at(14));
  assert.equal(a.pos('תיכון א', 0, 1).name, 'דנה לוי');
  assert.equal(a.pos('תיכון ב', 0, 1).name, 'עומר בן חיים');
  assert.equal(a.state.resetAt, at(5));
});
