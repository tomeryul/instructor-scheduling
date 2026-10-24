import { test } from 'node:test';
import assert from 'node:assert/strict';

import { createNameMatcher } from '../js/names.js';
import { parseCells } from '../js/schedule-parser.js';
import { parseRosterRows } from '../js/roster.js';
import { mergeSchedule, rankCandidates, autoFill, allAssignments, eventMatches } from '../js/scheduling.js';
import { buildDocument, toPlainText, slotLines, courseName } from '../js/export.js';

const rows = [
  ['שם', 'טלפון', 'הדרכה', 'הדרכת בטיחות', 'הערות'],
  ['דנה לוי', '1', 'V', 'V', 'קוסמת'],
  ['עומר בן חיים', '2', 'V', 'V', 'עושה הרצאות'],
  ['מאיה פרץ', '3', 'V', 'V', 'אילוצי סמסטר א\'- פנויה רק בחמישי עד 14:00'],
  ['רון אלמוג', '4', 'V', 'V', ''],
  ['תמר גל', '5', 'V', 'V', ''],
];
const instructors = parseRosterRows(rows, new Date('2026-10-24'));
const matcher = createNameMatcher(instructors.map(i => i.name));
const settings = {
  subjects: { chem: true, general: true, bio: false, cs: false, other: false },
  kinds: { study: true, course: true },
  slotKinds: { lab: true, lecture: true, magic: true },
  month: 'all',
};

const cells = [
  { date: '2026-11-02', lines: ['יום עיון בכימיה לתיכון א', '10:00-11:30 - מעבדה בכימיה. מדריכים: ?+?', '12:00-12:45 - מופע קסמים במדע. מדריך: ?'] },
  { date: '2026-11-02', lines: ['יום עיון בכימיה לתיכון ב', '10:30-12:00 - מעבדה בכימיה. מדריכים: דנה לוי + ?'] },
  { date: '2026-11-05', lines: ['יום עיון בביולוגיה לתיכון ג', '10:00-11:30 - מעבדה בביוטכנולוגיה. מדריכים: ?+?'] },
];

test('re-importing the schedule keeps what the user assigned', () => {
  const first = mergeSchedule([], parseCells(cells, matcher));
  const ev = first.events.find(e => e.title.includes('תיכון א'));
  ev.slots[0].groups[0].positions[0] = { name: 'רון אלמוג', source: 'user' };
  ev.ignored = false;
  ev.note = 'תוזכרו';

  // The coordinator edits the title a little and adds a day.
  const updated = [
    { ...cells[0], lines: ['יום עיון בכימיה לתיכון א (רכזת: 050)', ...cells[0].lines.slice(1)] },
    cells[1],
    { date: '2026-11-09', lines: ['יום עיון בכימיה לתיכון ד', '10:00-11:00 - מעבדה. מדריכים: ?+?'] },
  ];
  const second = mergeSchedule(first.events, parseCells(updated, matcher));
  const again = second.events.find(e => e.title.includes('תיכון א'));
  assert.equal(again.slots[0].groups[0].positions[0].name, 'רון אלמוג');
  assert.equal(again.note, 'תוזכרו');
  assert.equal(second.added, 1);
  // The biology day is outside the new file's dates (2-9 Nov)? No - it is inside, and untouched, so it goes.
  assert.equal(second.events.some(e => e.title.includes('תיכון ג')), false);
});

test('candidates: double-booked and unavailable people are not offered first', () => {
  const { events } = mergeSchedule([], parseCells(cells, matcher));
  const ev = events.find(e => e.title.includes('תיכון א'));
  const slot = ev.slots[0];
  const ranked = rankCandidates({ instructors, events, ev, slot, gi: 0, pi: 0 });
  const byName = Object.fromEntries(ranked.map(c => [c.instructor.name, c]));
  assert.equal(byName['דנה לוי'].tier, 'busy'); // already in תיכון ב at an overlapping time
  assert.equal(byName['מאיה פרץ'].tier, 'busy'); // Monday, only free on Thursday
  assert.equal(byName['רון אלמוג'].tier, 'good');
  const magic = rankCandidates({ instructors, events, ev, slot: ev.slots[1], gi: 0, pi: 0 });
  assert.equal(magic[0].instructor.name, 'דנה לוי');
});

test('auto-fill fills open positions without conflicts and skips hidden subjects', () => {
  const { events } = mergeSchedule([], parseCells(cells, matcher));
  const targets = events.filter(e => eventMatches(e, settings));
  const n = autoFill({ instructors, events, targets, settings });
  assert.equal(n, 4); // 2 + 1 (magic) + 1 in תיכון ב; the biology day is filtered out
  const bio = events.find(e => e.title.includes('תיכון ג'));
  assert.ok(bio.slots[0].groups[0].positions.every(p => !p.name));
  const all = allAssignments(events);
  for (const a of all) {
    const overlapping = all.filter(b => b !== a && b.key === a.key && b.ev.date === a.ev.date && b.start < a.end && b.end > a.start);
    assert.equal(overlapping.length, 0, `${a.name} is double-booked`);
  }
});

test('export: lines in the "משימות" format with ? for open positions', () => {
  const { events } = mergeSchedule([], parseCells(cells, matcher));
  const ev = events.find(e => e.title.includes('תיכון ב'));
  const text = slotLines(ev.slots[0]).map(r => r.map(x => x.text).join('')).join('\n');
  assert.equal(text, '10:30-12:00 - מעבדה בכימיה. מדריכים: דנה לוי + ?.');
  const doc = toPlainText(buildDocument({ events, instructors, settings, notes: { gaps: 'למצוא מחליף', tasks: '' } }, { month: 'all', today: '2026-11-01', sections: {} }));
  assert.match(doc, /פערים לטיפול:\n• למצוא מחליף/);
  assert.match(doc, /לא שובץ/);
  assert.match(doc, /יום שני- 02\/11-/);
  assert.doesNotMatch(doc, /ביולוגיה/);
  assert.equal(courseName("סביון שכבה ח (100 תל')- מפגש 1."), 'סביון שכבה ח');
});

// A schedule that changes from day to day: the user's choices and the coordinator's edits.
const school = 'יום עיון בכימיה לתיכון הדר רמלה';
const dayCells = (date, lab, magic) => [{ date, lines: [school, lab, magic] }];
const LAB = '10:30-12:00 - מעבדה בכימיה. מדריכים: ?+?';
const MAGIC = '12:30-13:15 - מופע קסמים. מדריך: ?';

function assignedDay() {
  const { events } = mergeSchedule([], parseCells(dayCells('2026-11-16', LAB, MAGIC), matcher));
  // As the picker does it: the position object is updated, not replaced.
  Object.assign(events[0].slots[0].groups[0].positions[0], { name: 'רון אלמוג', source: 'user' });
  Object.assign(events[0].slots[1].groups[0].positions[0], { name: 'דנה לוי', source: 'user' });
  return events;
}
const names = ev => ev.slots.map(s => s.groups.flatMap(g => g.positions.map(p => p.name)));

test('daily update: changed times keep the assignments and are reported', () => {
  const r = mergeSchedule(assignedDay(), parseCells(dayCells('2026-11-16', '10:00-11:30 - מעבדה בכימיה. מדריכים: ?+?', '12:00-12:45 - מופע קסמים. מדריך: ?'), matcher));
  assert.deepEqual(names(r.events[0]), [['רון אלמוג', null], ['דנה לוי']]);
  assert.equal(r.report.filter(x => x.level === 'attention').length, 2);
  assert.match(r.report[0].text, /10:30-12:00 ← 10:00-11:30/);
});

test('daily update: a name the coordinator wrote replaces the user\'s and is reported', () => {
  const r = mergeSchedule(assignedDay(), parseCells(dayCells('2026-11-16', '10:30-12:00 - מעבדה בכימיה. מדריכים: תמר גל+?', MAGIC), matcher));
  assert.deepEqual(names(r.events[0]), [['תמר גל', null], ['דנה לוי']]);
  assert.match(r.report.find(x => x.level === 'attention').text, /תמר גל במקום רון אלמוג/);
  // The next update with the same file keeps it quiet.
  const again = mergeSchedule(r.events, parseCells(dayCells('2026-11-16', '10:30-12:00 - מעבדה בכימיה. מדריכים: תמר גל+?', MAGIC), matcher));
  assert.equal(again.report.length, 0);
});

test('daily update: the user\'s choice stands while the file says the same as before', () => {
  const events = mergeSchedule([], parseCells(dayCells('2026-11-16', '10:30-12:00 - מעבדה בכימיה. מדריכים: תמר גל+?', MAGIC), matcher)).events;
  Object.assign(events[0].slots[0].groups[0].positions[0], { name: 'רון אלמוג', source: 'user' });
  const r = mergeSchedule(events, parseCells(dayCells('2026-11-16', '10:30-12:00 - מעבדה בכימיה. מדריכים: תמר גל+?', MAGIC), matcher));
  assert.equal(r.events[0].slots[0].groups[0].positions[0].name, 'רון אלמוג');
});

test('daily update: a name removed from the file is not put back', () => {
  const events = mergeSchedule([], parseCells(dayCells('2026-11-16', '10:30-12:00 - מעבדה בכימיה. מדריכים: תמר גל+?', MAGIC), matcher)).events;
  const r = mergeSchedule(events, parseCells(dayCells('2026-11-16', LAB, MAGIC), matcher));
  assert.deepEqual(names(r.events[0]), [[null, null], [null]]);
  assert.match(r.report[0].text, /תמר גל הוסר/);
});

test('daily update: a study day moved to another date takes its assignments along', () => {
  const r = mergeSchedule(assignedDay(), parseCells([
    ...dayCells('2026-11-18', LAB, MAGIC),
    { date: '2026-11-16', lines: ['יום עיון בכימיה לתיכון אחר', '10:00-11:00 - מעבדה. מדריכים: ?+?'] },
  ], matcher));
  const moved = r.events.find(e => e.date === '2026-11-18');
  assert.deepEqual(names(moved), [['רון אלמוג', null], ['דנה לוי']]);
  assert.match(moved.changes[0].text, /הוזז מ-16\/11 ל-18\/11/);
  const other = r.events.find(e => e.date === '2026-11-16');
  assert.deepEqual(names(other), [[null, null]]);
  assert.equal(other.changes[0].level, 'new');
});

test('daily update: a cancelled day with assignments stays, flagged', () => {
  const r = mergeSchedule(assignedDay(), parseCells([{ date: '2026-11-17', lines: ['יום עיון בכימיה לתיכון אחר', '10:00-11:00 - מעבדה. מדריכים: ?+?'] }, { date: '2026-11-15', lines: ['יום עיון בכימיה לתיכון שלישי', '10:00-11:00 - מעבדה. מדריכים: ?+?'] }], matcher));
  const gone = r.events.find(e => e.date === '2026-11-16');
  assert.equal(gone.missingFromSchedule, true);
  assert.match(gone.changes[0].text, /רון אלמוג, דנה לוי/);
});
