// Example data shown before the user imports their own files. Every name, school
// and phone number here is made up; the cells go through the same parser as a
// real schedule so the example behaves like the real thing.

import { parseCells } from './schedule-parser.js';
import { parseRosterRows } from './roster.js';
import { createNameMatcher } from './names.js';
import { toAppEvent } from './scheduling.js';

const ROSTER = [
  ['שם מדריך', 'מספר טלפון', 'הדרכה', 'הכשרה ראשונה', 'קבוצת ווצפ', 'תואר ראשון', 'הדרכת בטיחות', 'הערות'],
  ['דנה לוי', '050-0000001', 'V', '', 'V', 'כימיה', 'V', 'קוסמת\nאילוצי סמסטר א\'- לא יכולה ימי שלישי'],
  ['עומר בן חיים', '050-0000002', 'V', '', 'V', 'כימיה', 'V', 'עושה הרצאות'],
  ['מאיה פרץ', '050-0000003', 'V', '', 'V', 'כימיה', 'V', 'אילוצי סמסטר א\'- פנויה רק בחמישי עד 14:00'],
  ['יואב שמעוני', '050-0000004', 'V', '', 'V', 'כימיה', 'V', 'קוסם'],
  ['נועה אדלר', '050-0000005', 'V', '', 'V', 'ביוטכנולוגיה', 'V', 'אילוצי סמסטר א\'- בשני יכולה החל משעה 12:00'],
  ['אביב רוזן', '050-0000006', 'V', '', 'V', 'כימיה', '05/10- לא השלים, בטיפול', 'חדש'],
  ['תמר גל', '050-0000007', 'V', '', 'V', 'כימיה', 'V', 'הכי נוח לה ימי ראשון וחמישי'],
  ['רון אלמוג', '050-0000008', 'V', '', 'V', 'כימיה', 'V', ''],
  ['שירן כץ', '050-0000009', '', '', 'X- אין צורך', '', '', 'עושה הרצאות'],
];

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export function demoState(today = new Date()) {
  const sunday = new Date(today);
  sunday.setDate(sunday.getDate() + ((7 - sunday.getDay()) % 7 || 7));
  const day = n => addDays(sunday, n);
  const cells = [
    { date: day(0), color: '#ffff00', lines: [
      'יום עיון בכימיה לתיכון לדוגמה א (רכזת: 050-1111111) שתי קבוצות - כיתה י"א - 40 תל\'. 5 ש"א.',
      '9:45-10:00- התכנסות ברחבת הספרייה',
      '10:00-10:45- הרצאה בנושא: פולימרים. מרצה: ?.',
      '11:00-13:15 - מעבדה בכימיה- גזים ואלקטרוליזה. מדריכים: ?+?. מיקום: מעבדה 104.',
      '13:15-13:45 - הפסקה',
      '13:45-14:30 - מופע קסמים במדע. מדריך: ?.',
    ] },
    { date: day(1), color: '#ffff00', lines: [
      'יום עיון בכימיה לחטיבת ביניים לדוגמה (רכז: 050-2222222) כיתה ח\' - 30 תל\'. 4 ש"א.',
      '10:00-11:30- חומציות של קרח יבש. מדריכים: דנה לוי + ?',
      '12:00-13:30 - פצצת אמבטיה. מדריכים: דנה לוי + ?',
    ] },
    { date: day(2), color: '#ffc000', lines: [
      'קורס כימיה לדוגמה - מפגש 3/7',
      '10:00-11:30 -',
      'קבוצה א\'- קורס כימיה- בישול מולקולרי. מדריכים: ?+?.',
      'קבוצה ב\'- קורס כימיה- בישול מולקולרי. מדריכים: ?+?.',
    ] },
    { date: day(4), color: '#ffff00', lines: [
      'יום עיון לתיכון לדוגמה ב (רכזת: 050-3333333) כיתה י"ב - 28 תל\', 4 ש"א.',
      '10:00-10:45 – הרצאה. מרצה: עומר בן חיים',
      '11:00-12:30 – ניסוי חקר בכימיה - הפקת גזים. מדריכים: ?+?',
      '13:00-13:45- מופע קסמים במדע- ?',
    ] },
    { date: day(7), color: '#a6c9ec', lines: [
      'תוכנית ערב לדוגמה - מפגש 2',
      'שעות: 17:00-18:30',
      'מדריך: מישהו מתוכנית אחרת',
    ] },
  ];
  const instructors = parseRosterRows(ROSTER, today);
  const events = parseCells(cells, createNameMatcher(instructors.map(i => i.name))).map(toAppEvent);
  return { instructors, events };
}
