// Turns the text of a schedule cell into an event with staffing slots:
// a study day / course, its time slots (lab, lecture, magic show) and the
// instructor positions in each slot, with names already written in the cell.

import { cleanName } from './names.js';

export const EVENT_KINDS = { study: 'יום עיון', course: 'קורס', other: 'אחר' };
export const SLOT_KINDS = { lab: 'הדרכה', lecture: 'הרצאה', magic: 'מופע קסמים' };
export const SUBJECTS = { chem: 'כימיה', bio: 'ביולוגיה', cs: 'מחשבים', other: 'אחר', general: 'כללי' };

const TIME_RANGE_RE = /(\d{1,2})[:.](\d{2})\s*[-–—]\s*(\d{1,2})[:.](\d{2})/g;
const ROLE_RE = /(מדריכים|מדריכות|מדריכה|מדריך|מרצים|מרצות|מרצה|מפי)(?![א-ת])(?!\s*(?:חברתי|מלווים|מלווה))\s*[:\-–]?\s*/g;
const STAFF_END_RE = /\.(?:\s|$)|\s(?:מיקום|בנושא|נושא|בשעה|שעות)(?![א-ת])|\s(?:מדריכ|מרצ)|\n|$/;
const SKIP_SLOT_RE = /^(?:התכנסות|הפסק|הגעה|מעבר|טקס|סיכום|דברי|ארוחה|פיזור|כיבוד|פעילות חברתית|הזמנת)/;
const OTHER_EVENT_RE = /תוכנית\s+(?:ספיר|שוהם|ענבר|עבבר)|מפגש היערכות|מכתבי קבלה|מפגש פתיחה|^תוכנית(?![^\n]*(?:גלילאו|קוונטום))/;
const STUDY_RE = /יום עיון|יום שיא/;
const COURSE_RE = /קורס|גלילאו|סביון|קוונטום|אורט נופים|מפגש\s*\d+\s*(?:\/|מתוך)\s*\d+/;
const GROUP_RE = /קבוצה\s+([א-ת])(?:['׳]|(?=[\s:\-–]))\s*[:\-–]?\s*/g;
// A course heading starts a line (or follows the closing bracket of the previous course's text).
const COURSE_HEADING_RE = /(?:^|\n|\)\s)קורס\s+(?!בחירה|\d)([^\n\-–:.]{2,40}?)(?=\s*(?:\n|מיקום|$))/g;
const TWO_GROUPS_RE = /(?:שתי|2)\s+קבוצות(?!\s+אחת אחרי)/;

// Explicit subject words win over activity keywords.
const SUBJECT_RULES = [
  ['cs', /מיינקראפט|פייתון|בינה מלאכותית|מחשבים|תכנות|אלגוריתמ/],
  ['bio', /ביולוג|ביוטכנולוג|זאולוג/],
  ['chem', /כימי/],
  ['other', /צרפתית|קרימינולוג|פסיכולוג|משפט|מחוננים/],
  ['bio', /(?:^|[^א-ת])דנ"?א(?![א-ת])|חיידק|מיקרוסקופ|צמחים|ויטמין|elegans|הנדסה גנטית|גראם/],
  ['chem', /גזים|אלקטרוליז|לומינול|בישול מולקולרי|פולימר|קרח יבש|מראת הכסף|חמצון|צביעת בדים|פצצת אמבטיה|כרומטוגרפ|עוגיות|גביש|צלליות|קסמים/],
];

export function detectSubject(text) {
  for (const [subject, re] of SUBJECT_RULES) if (re.test(text)) return subject;
  return null;
}

export function toMinutes(t) {
  const [h, m] = String(t).split(':').map(Number);
  return h * 60 + m;
}

const pad = n => String(n).padStart(2, '0');
const fmt = (h, m) => `${pad(+h)}:${m}`;

/** All time ranges in a text, each as {start, end, index, length}; reversed ranges are put right. */
export function findTimeRanges(text) {
  const out = [];
  for (const m of text.matchAll(TIME_RANGE_RE)) {
    let start = fmt(m[1], m[2]), end = fmt(m[3], m[4]);
    if (toMinutes(start) > toMinutes(end)) [start, end] = [end, start];
    out.push({ start, end, index: m.index, length: m[0].length });
  }
  return out;
}

function hashString(s) {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619);
  return (h >>> 0).toString(36);
}

const tidy = s => s
  .replace(/\s+/g, ' ')
  .replace(/(?:\s*[.,]){2,}/g, '.')
  .replace(/\s+([.,])/g, '$1')
  .replace(/^[\s\-–—:.,]+|[\s\-–—:,.]+$/g, '')
  .trim();

/** "מדריכים: א + ב, ג + ד" -> [[א, ב], [ג, ד]] with null for every "?". */
function parseStaffList(text, matcher) {
  const groups = [];
  for (const part of text.split(/\s*[,;]\s*/)) {
    if (!part.trim()) continue;
    const positions = [];
    for (const token of part.split(/\s*\+\s*/)) {
      const t = token.trim();
      if (!t) continue;
      if (/^[?_\s]+$/.test(t)) { positions.push(null); continue; }
      const hits = matcher.findAll(t);
      if (hits.length) hits.forEach(h => positions.push(h.name));
      else if (/[א-תA-Za-z]/.test(t)) positions.push(cleanName(t));
      if (/\?/.test(t) && !hits.length && !/[א-תA-Za-z]/.test(t)) positions.push(null);
    }
    if (positions.length) groups.push(positions);
  }
  return groups;
}

/**
 * Staff written in a piece of slot text.
 * @returns {{groups: (string|null)[][], rest: string}} rest is the text without the staff part.
 */
function extractStaff(text, matcher) {
  const groups = [];
  let rest = text;
  ROLE_RE.lastIndex = 0;
  const markers = [...text.matchAll(ROLE_RE)];
  if (markers.length) {
    // Work from the last marker backwards so the indexes stay valid.
    for (const m of markers.reverse()) {
      const from = m.index + m[0].length;
      const tail = text.slice(from);
      const endMatch = tail.match(STAFF_END_RE);
      const staffText = tail.slice(0, endMatch ? endMatch.index : tail.length);
      groups.unshift(...parseStaffList(staffText, matcher));
      rest = rest.slice(0, m.index) + ' ' + rest.slice(from + staffText.length);
    }
  } else {
    const hits = matcher.findAll(text);
    if (hits.length) groups.push(hits.map(h => h.name));
    for (const h of [...hits].reverse()) rest = rest.slice(0, h.from) + ' ' + rest.slice(h.to);
  }
  return { groups, rest };
}

function extractLocation(text) {
  const m = text.match(/מיקום\s*:?\s*([^.,\n]*)/);
  const location = m ? tidy(m[1]).replace(/^\?$/, '') : '';
  return { location, rest: m ? text.replace(m[0], ' ') : text };
}

function makeGroups(found, kind) {
  if (found.length) return found.map(positions => ({ label: '', positions }));
  return [{ label: '', positions: kind === 'lab' ? [null, null] : [null] }];
}

function slotKind(body) {
  if (/קסמים/.test(body)) return 'magic';
  if (/^הרצא|הרצאה|מרצה|מפי\s/.test(body)) return 'lecture';
  if (/מעבד|ניסוי|בישול|סדנ|קורס|פעילות|חקר|סיור|מדריכ|\?\s*\+\s*\?/.test(body)) return 'lab';
  return null;
}

/** One time slot: label, location, kind and its groups of positions. */
function parseSlot(range, body, matcher) {
  const clean = body.replace(/^[\s\-–—:]+/, '');
  if (SKIP_SLOT_RE.test(clean)) return null;
  const kind = slotKind(clean);
  if (!kind) return null;

  // "קבוצה א'- ... קבוצה ב'- ..." splits the slot into labelled groups.
  GROUP_RE.lastIndex = 0;
  const groupMarks = [...clean.matchAll(GROUP_RE)];
  const head = groupMarks.length ? clean.slice(0, groupMarks[0].index) : clean;
  const { location, rest: noLocation } = extractLocation(head);
  let groups = [];
  let labelText = noLocation;
  if (groupMarks.length) {
    groupMarks.forEach((g, i) => {
      const chunk = clean.slice(g.index + g[0].length, i + 1 < groupMarks.length ? groupMarks[i + 1].index : clean.length);
      const { groups: found, rest } = extractStaff(chunk, matcher);
      const activity = tidy(extractLocation(rest).rest.replace(/מעבדה\s+\d+\s+בניין\s+\d+\.?/g, ''));
      const positions = found.flat();
      groups.push({
        label: `קבוצה ${g[1]}'`,
        activity,
        subject: detectSubject(chunk),
        positions: positions.length ? positions : kind === 'lab' ? [null, null] : [null],
      });
    });
    labelText = extractStaff(noLocation, matcher).rest;
  } else {
    const { groups: found, rest } = extractStaff(noLocation, matcher);
    groups = makeGroups(found, kind);
    labelText = rest;
  }

  let label = tidy(labelText);
  if (!label) label = [...new Set(groups.map(g => g.activity).filter(Boolean))].join(' / ');
  if (!label) label = SLOT_KINDS[kind];
  return { start: range.start, end: range.end, kind, label, location, groups, subject: detectSubject(clean) };
}

/** Groups of different subjects in one slot (course A for group 1, course B for group 2) become separate slots. */
function splitBySubject(slot) {
  const subjects = [...new Set(slot.groups.map(g => g.subject || null))];
  if (subjects.length < 2) return [slot];
  return subjects.map(subject => {
    const groups = slot.groups.filter(g => (g.subject || null) === subject);
    const activity = groups.find(g => g.activity)?.activity;
    return { ...slot, subject: subject || slot.subject, label: activity || slot.label, groups };
  });
}

/** A cell listing several courses (סביון): one slot per course, all at the course hours. */
function parseCourseBlocks(text, matcher) {
  COURSE_HEADING_RE.lastIndex = 0;
  const headings = [...text.matchAll(COURSE_HEADING_RE)].map(m => {
    const at = m.index + m[0].indexOf('קורס');
    return { index: at, end: m.index + m[0].length, name: m[1] };
  });
  if (headings.length < 2) return null;
  const ranges = findTimeRanges(text.slice(0, headings[0].index));
  const courseRanges = ranges.filter(r => /קורס/.test(text.slice(r.index + r.length, r.index + r.length + 25)));
  const span = courseRanges.length ? courseRanges : ranges;
  if (!span.length) return null;
  const start = span.map(r => r.start).sort()[0];
  const end = span.map(r => r.end).sort().pop();
  return headings.map((h, i) => {
    const block = text.slice(h.end, i + 1 < headings.length ? headings[i + 1].index : text.length);
    const name = tidy(h.name);
    const { location, rest } = extractLocation(block);
    const { groups: found } = extractStaff(rest, matcher);
    const activity = (block.match(/פעילות\s+([^(\n]*)/) || [])[1];
    const label = `קורס ${name}${activity && tidy(activity) && tidy(activity) !== '?' ? ' - ' + tidy(activity) : ''}`;
    const positions = found.flat();
    return {
      start, end, kind: 'lab', label, location,
      subject: detectSubject(`קורס ${name}`) || detectSubject(block),
      groups: [{ label: '', positions: positions.length ? positions : [null, null] }],
    };
  });
}

function eventKind(text, slots) {
  if (OTHER_EVENT_RE.test(text)) return 'other';
  if (STUDY_RE.test(text)) return 'study';
  if (COURSE_RE.test(text)) return 'course';
  return slots.length ? 'study' : 'other';
}

/**
 * @param cell     {date, lines: string[], color}
 * @param matcher  name matcher from createNameMatcher
 * @returns event  {id, date, color, kind, subject, title, raw, slots}
 */
export function parseCell(cell, matcher) {
  const raw = cell.lines.join('\n');
  const ranges = findTimeRanges(raw);
  const title = tidy((ranges.length ? raw.slice(0, ranges[0].index) : raw).replace(/\n/g, ' '));
  const headerSubject = detectSubject(title.replace(/\(.*?\)/g, ''));

  let slots = parseCourseBlocks(raw, matcher);
  if (!slots) {
    slots = [];
    ranges.forEach((r, i) => {
      const body = raw.slice(r.index + r.length, i + 1 < ranges.length ? ranges[i + 1].index : raw.length).replace(/\n/g, ' ');
      // A bare "המשך" just extends the slot before it.
      if (/^[\s\-–—:]*המשך[\s.]*$/.test(body) && slots.length) { slots[slots.length - 1].end = r.end; return; }
      const slot = parseSlot(r, body, matcher);
      if (slot) slots.push(...splitBySubject(slot));
    });
  }

  // "שתי קבוצות" in the header and a lab with a single empty pair: one pair per group.
  if (TWO_GROUPS_RE.test(title)) {
    for (const s of slots) {
      if (s.kind === 'lab' && s.groups.length === 1 && !s.groups[0].label && s.groups[0].positions.every(p => !p)) {
        s.groups = [{ label: '', positions: [null, null] }, { label: '', positions: [null, null] }];
      }
    }
  }

  const kind = eventKind(raw, slots);
  const id = 'e' + hashString(cell.date + '|' + title.slice(0, 60));
  slots.forEach((s, i) => {
    s.id = `${id}.${hashString(`${s.start}-${s.end}|${s.kind}|${s.label}`)}${i}`;
    s.subject = s.subject || headerSubject || null;
    s.groups.forEach(g => { g.positions = g.positions.map(p => p || null); g.activity = g.activity || ''; delete g.subject; });
  });
  return { id, date: cell.date, color: cell.color || null, kind, subject: headerSubject, title, raw, slots };
}

/** Parse all cells; events that share an id (same title on the same day) get a suffix. */
export function parseCells(cells, matcher) {
  const seen = new Map();
  return cells.map(cell => {
    const ev = parseCell(cell, matcher);
    const n = seen.get(ev.id) || 0;
    seen.set(ev.id, n + 1);
    if (n) {
      const oldId = ev.id;
      ev.id = `${oldId}_${n}`;
      ev.slots.forEach(s => { s.id = s.id.replace(oldId, ev.id); });
    }
    return ev;
  });
}
