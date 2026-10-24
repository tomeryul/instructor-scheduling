import { test } from 'node:test';
import assert from 'node:assert/strict';

import { normalizeName, createNameMatcher } from '../js/names.js';
import { parseCell, parseCells, findTimeRanges, detectSubject } from '../js/schedule-parser.js';
import { parseAvailability, parseRosterRows, checkAvailability, parseWindows } from '../js/roster.js';
import { lineToText } from '../js/pdf-grid.js';

const ROSTER = ['דנה לוי', 'עומר בן חיים', 'מאיה שא-נס פרץ', 'היבא כהן', 'רון בוצקו'];
const matcher = createNameMatcher(ROSTER);

test('names: spelling variants match the roster', () => {
  assert.equal(normalizeName("בוצ'קו"), normalizeName('בוצקו'));
  assert.equal(matcher.match('היבה כהן'), 'היבא כהן');
  assert.equal(matcher.match("רון בוצ'קו"), 'רון בוצקו');
  assert.equal(matcher.match('מאיה פרץ'), 'מאיה שא-נס פרץ');
  assert.equal(matcher.match('דנה'), null);
  const hits = matcher.findAll('מדריכים: דנה לוי + עומר בן חיים.');
  assert.deepEqual(hits.map(h => h.name), ['דנה לוי', 'עומר בן חיים']);
});

test('time ranges: written backwards are put right', () => {
  const r = findTimeRanges('12:15 – 10:00 - מעבדה, 9:45-10:00 התכנסות');
  assert.deepEqual(r.map(x => [x.start, x.end]), [['10:00', '12:15'], ['09:45', '10:00']]);
});

test('study day: lecture, lab with two groups, magic show; breaks skipped', () => {
  const ev = parseCell({
    date: '2026-11-16',
    lines: [
      'יום עיון בכימיה לתיכון לדוגמה (רכזת: 050-0000000) כיתה י"א - 40 תל\'. 5 ש"א.',
      '10:30 – 10:15 - התכנסות ברחבה',
      '11:15 – 10:30 - הרצאה בנושא תכשיטים מדריך: רון בוצ\'קו. מיקום:?.',
      '11:30 – 11:15 - הפסקה',
      '13:45 – 11:30 - ניסוי גזים ואלקטרוליזה. קבוצה א\'- מדריכים: דנה לוי + עומר בן חיים. מיקום: ?.',
      'קבוצה ב\'- מדריכים: ?+?. מיקום: ?.',
      '15:00 – 14:15 - מופע קסמים במדע- היבה כהן. מיקום: מעבדה 105 בניין 202.',
    ],
  }, matcher);
  assert.equal(ev.kind, 'study');
  assert.equal(ev.subject, 'chem');
  assert.deepEqual(ev.slots.map(s => [s.start, s.end, s.kind]), [
    ['10:30', '11:15', 'lecture'], ['11:30', '13:45', 'lab'], ['14:15', '15:00', 'magic'],
  ]);
  assert.deepEqual(ev.slots[0].groups[0].positions, ['רון בוצקו']);
  assert.deepEqual(ev.slots[1].groups.map(g => [g.label, g.positions]), [
    ["קבוצה א'", ['דנה לוי', 'עומר בן חיים']], ["קבוצה ב'", [null, null]],
  ]);
  assert.equal(ev.slots[1].label, 'ניסוי גזים ואלקטרוליזה');
  assert.deepEqual(ev.slots[2].groups[0].positions, ['היבא כהן']);
  assert.equal(ev.slots[2].label, 'מופע קסמים במדע');
  assert.equal(ev.slots[2].location, 'מעבדה 105 בניין 202');
});

test('study day: "שתי קבוצות" with one empty pair becomes two pairs', () => {
  const ev = parseCell({ date: '2026-11-22', lines: ['יום עיון בכימיה לתיכון לדוגמה שתי קבוצות - כיתה י"א.', '11:00-13:15 - מעבדה בכימיה. מדריכים: ?+?.'] }, matcher);
  assert.equal(ev.slots[0].groups.length, 2);
  assert.deepEqual(ev.slots[0].groups.map(g => g.positions.length), [2, 2]);
});

test('a cell with several courses gets one slot per course, at the course hours', () => {
  const ev = parseCell({
    date: '2026-11-05',
    lines: [
      'סביון שכבה ח (100 תל\')- מפגש 1.',
      '09:30-10:00 התכנסות.',
      '10:00-11:30 קורס בחירה 1',
      '12:00-13:30 קורס בחירה 2',
      'מדריכים מלווים: ?',
      'קורס כימיה',
      'מיקום: בניין 202 קומה 1 מעבדה 104, מדריכים: דנה לוי + עומר בן חיים פעילות תעלומת העוגיות (2 קבוצות אחת אחרי השנייה) קורס ביוטכנולוגיה',
      'מיקום: בניין 202 מעבדה 105, מדריכים: ?+?',
    ],
  }, matcher);
  assert.equal(ev.kind, 'course');
  assert.equal(ev.slots.length, 2);
  assert.deepEqual([ev.slots[0].start, ev.slots[0].end], ['10:00', '13:30']);
  assert.equal(ev.slots[0].subject, 'chem');
  assert.equal(ev.slots[0].label, 'קורס כימיה - תעלומת העוגיות');
  assert.deepEqual(ev.slots[0].groups[0].positions, ['דנה לוי', 'עומר בן חיים']);
  assert.equal(ev.slots[1].subject, 'bio');
});

test('groups of different courses in one time slot are split by subject', () => {
  const ev = parseCell({
    date: '2026-11-17',
    lines: ['אורט לדוגמה - מפגש 2/7', '10:00-11:30 -', "קבוצה א'- קורס כימיה- גזים. מדריכים: דנה לוי+עומר בן חיים.", "קבוצה ב'- קורס ביוטכנולוגיה- ?. מדריכים: ?+?."],
  }, matcher);
  assert.deepEqual(ev.slots.map(s => s.subject), ['chem', 'bio']);
  assert.equal(ev.slots[0].label, 'קורס כימיה- גזים');
});

test('programs and planning meetings are not staffing events', () => {
  const ev = parseCell({ date: '2026-11-01', lines: ['תוכנית ספיר מחזור ג', 'שעות: 17:00-18:30', 'מדריך: מישהו'] }, matcher);
  assert.equal(ev.kind, 'other');
});

test('subjects: "סדנא" is not DNA', () => {
  assert.equal(detectSubject('מעבדת סדנא צלליות במלח'), 'chem');
  assert.equal(detectSubject('הפקת דנ"א'), 'bio');
});

test('same title on the same day gets distinct ids', () => {
  const cell = { date: '2026-11-03', lines: ['יום עיון בכימיה', '10:00-11:00 מעבדה. מדריכים: ?+?'] };
  const [a, b] = parseCells([cell, cell], matcher);
  assert.notEqual(a.id, b.id);
  assert.notEqual(a.slots[0].id, b.slots[0].id);
});

test('availability: restrictions leave other days free', () => {
  const a = parseAvailability("אילוצי סמסטר א'- לא יכול בראשון, בשני יכול החל משעה 12:00 וברביעי יכול עד שעה 15:00", new Date('2026-10-24'));
  assert.equal(a.days[0], 'no');
  assert.deepEqual(a.days[1], [['12:00', '24:00']]);
  assert.deepEqual(a.days[3], [['00:00', '15:00']]);
  assert.equal(a.days[2], undefined);
});

test('availability: listing free times closes the other weekdays', () => {
  const a = parseAvailability("תואר ראשון\nאילוצי סמסטר א'- פנויה ב:\nראשון אחרי 12:00\nשני אחרי 12:00\nשלישי עד 13:00\nחמישי עד 13:00");
  assert.deepEqual(a.days[0], [['12:00', '24:00']]);
  assert.equal(a.days[3], 'no');
  const b = parseAvailability('שלישי- עד 11:00 ומ-12:00 פנוי\nחמישי- לא פנוי');
  assert.deepEqual(b.days[2], [['00:00', '11:00'], ['12:00', '24:00']]);
  assert.equal(b.days[4], 'no');
});

test('availability: shared wording, preferences and reserve duty', () => {
  assert.deepEqual(Object.keys(parseAvailability('מלמדת בשני רביעי וחמישי').days), ['1', '3', '4']);
  const pref = parseAvailability('עדיפות לימי א,ב,ד.\nשעות 9-14.');
  assert.deepEqual(pref.preferredDays, [0, 1, 3]);
  assert.deepEqual(pref.preferredHours, ['09:00', '14:00']);
  assert.deepEqual(parseAvailability('הכי נוח לה ימי ראשון וחמישי').preferredDays, [0, 4]);
  assert.equal(parseAvailability('במילואים עד סוף דצמבר', new Date('2026-10-24')).unavailableUntil, '2026-12-31');
  assert.deepEqual(parseWindows('יום שני 10-13.45'), [['10:00', '13:45']]);
  assert.deepEqual(parseWindows('14 והלאה אפשרי'), [['14:00', '24:00']]);
});

test('roster rows: roles from the guide column and the notes', () => {
  const rows = [
    ['שם מדריך', 'מספר טלפון', 'הדרכה', 'קבוצת ווצפ', 'תואר ראשון', 'הדרכת בטיחות', 'הערות'],
    ['א ב', '050', 'V', 'V', 'כימיה', 'V', 'קוסם\nאילוצי סמסטר א\'- לא יכול ימי שלישי'],
    ['ג ד', '050', 'X', 'V', 'כימיה', 'V', 'רק קוסם, בערבית'],
    ['ה ו', '050', '', 'X- אין צורך', '', '', 'עושה הרצאות'],
    ['ז ח', '050', 'V', 'V', '', 'V', 'כבר לא רוצה להדריך'],
  ];
  const [a, b, c, d] = parseRosterRows(rows);
  assert.deepEqual(a.roles, { lab: true, magic: true, lecture: false });
  assert.equal(checkAvailability(a, '2026-11-03', '10:00', '11:00').level, 'no'); // a Tuesday
  assert.deepEqual(b.roles, { lab: false, magic: true, lecture: false });
  assert.ok(b.tags.includes('ערבית'));
  assert.deepEqual(c.roles, { lab: false, magic: false, lecture: true });
  assert.equal(d.active, false);
});

test('pdf lines: right-to-left order with phone numbers kept left-to-right', () => {
  // Visual items of "(רכזת: 050-1234567) כיתה י" in a right-to-left line: the number
  // block sits left of the colon and reads left to right inside.
  const items = [
    { str: '(', x0: 99, x1: 100 }, { str: 'רכזת', x0: 95, x1: 98.9 }, { str: ':', x0: 94.6, x1: 94.9 },
    { str: '1234567', x0: 89, x1: 94.4 }, { str: '-', x0: 88.6, x1: 89 }, { str: '050', x0: 86.2, x1: 88.6 },
    { str: ')', x0: 85.8, x1: 86.2 }, { str: 'כיתה י', x0: 79, x1: 85.3 },
  ];
  assert.equal(lineToText(items, 1.1), '(רכזת: 050-1234567) כיתה י');
});
