// The instructor list ("נוער שוחר מדע - שיבוץ מדריכים"): who can guide labs, who is a
// magician or a lecturer, safety-training status, and the free-text availability
// notes turned into weekly rules.

import { normalizeName } from './names.js';
import { toMinutes } from './schedule-parser.js';

export const DAY_NAMES = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];
const DAY_LETTERS = { 'א': 0, 'ב': 1, 'ג': 2, 'ד': 3, 'ה': 4, 'ו': 5 };

const COLUMNS = {
  name: /שם/,
  phone: /טלפון|נייד|פלאפון/,
  guide: /^הדרכה$|^מדריך/,
  whatsapp: /ווצ|וואטס|whats/i,
  degree: /תואר|מקצוע/,
  safety: /בטיחות/,
  notes: /הערות/,
};

const yes = v => /^\s*v\s*$/i.test(String(v)) || /^\s*(כן|✓|✔)\s*$/.test(String(v));
const no = v => /^\s*x\b/i.test(String(v)) || /^\s*לא\s*$/.test(String(v));

const pad = n => String(n).padStart(2, '0');
function hhmm(h, m) {
  const hours = +h;
  return hours >= 24 ? '24:00' : `${pad(hours)}:${pad(+(m || 0))}`;
}

const MONTHS = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];

/** Windows ([from, to] in HH:MM) written in a piece of text: "10-13.45", "מ-12:00", "עד 14:00", "14 והלאה". */
export function parseWindows(text) {
  const windows = [];
  let rest = text;
  for (const m of text.matchAll(/(\d{1,2})(?:[:.](\d{2}))?\s*[-–]\s*(\d{1,2})(?:[:.](\d{2}))?(?!\d)/g)) {
    windows.push([hhmm(m[1], m[2]), hhmm(m[3], m[4])]);
    rest = rest.replace(m[0], ' ');
  }
  const tokens = [];
  const time = '(\\d{1,2})(?:[:.](\\d{2}))?';
  const re = new RegExp(`(?:(?:החל\\s+)?(?<![א-ת])ו?מ(?:שעה)?\\s*-?\\s*|אחרי\\s+(?:שעה\\s+)?)${time}|${time}\\s*(?:והלאה|ואילך)|עד\\s*(?:שעה\\s*)?${time}`, 'g');
  for (const m of rest.matchAll(re)) {
    if (m[1] !== undefined) tokens.push({ type: 'from', t: hhmm(m[1], m[2]) });
    else if (m[3] !== undefined) tokens.push({ type: 'from', t: hhmm(m[3], m[4]) });
    else tokens.push({ type: 'to', t: hhmm(m[5], m[6]) });
  }
  let open = null;
  for (const tk of tokens) {
    if (tk.type === 'from') {
      if (open) windows.push([open, '24:00']);
      open = tk.t;
    } else {
      windows.push([open || '00:00', tk.t]);
      open = null;
    }
  }
  if (open) windows.push([open, '24:00']);
  return windows;
}

/**
 * Free-text constraints -> weekly rules.
 * @returns {{days: Object<number, 'no'|[string,string][]>, preferredDays: number[], preferredHours: [string,string]|null,
 *            unavailableUntil: string|null, recognized: boolean}}
 */
export function parseAvailability(notes, today = new Date()) {
  const result = { days: {}, preferredDays: [], preferredHours: null, unavailableUntil: null, recognized: false };
  if (!notes) return result;
  const text = String(notes)
    .replace(/תואר\s+ראשון/g, ' ')
    .replace(/אילוצי\s+סמסטר\s+[א-ת]['׳]?\s*[-–:]?/g, '\n')
    .replace(/סמסטר\s+[א-ת]['׳]?/g, ' ');

  // "במילואים עד סוף דצמבר"
  const until = text.match(/עד\s+(?:סוף\s+)?(?:חודש\s+)?(ינואר|פברואר|מרץ|אפריל|מאי|יוני|יולי|אוגוסט|ספטמבר|אוקטובר|נובמבר|דצמבר)/);
  if (until && /מילואים|לא זמינ|לא פנוי|חופש|חו"ל|בחו/.test(text)) {
    const month = MONTHS.indexOf(until[1]);
    const year = month >= today.getMonth() ? today.getFullYear() : today.getFullYear() + 1;
    const last = new Date(year, month + 1, 0).getDate();
    result.unavailableUntil = `${year}-${pad(month + 1)}-${pad(last)}`;
    result.recognized = true;
  }

  const dayRe = /(?<![א-ת])[ובה]{0,2}(ראשון|שני|שלישי|רביעי|חמישי|שישי)(?![א-ת])|ימי\s+((?:[א-ו]['׳]?\s*[,ו]?\s*){2,})/g;
  let hasNegative = false;
  let hasListing = false;
  let lastWasPreference = false;

  for (const segment of text.split(/\n|;|\.(?=\s|$)/)) {
    const mentions = [];
    for (const m of segment.matchAll(dayRe)) {
      const days = m[1] ? [DAY_NAMES.indexOf(m[1])] : [...m[2].matchAll(/[א-ו]/g)].map(l => DAY_LETTERS[l[0]]);
      mentions.push({ days, index: m.index, end: m.index + m[0].length });
    }
    if (!mentions.length) {
      const windows = parseWindows(segment);
      if (windows.length && lastWasPreference) { result.preferredHours = windows[0]; result.recognized = true; }
      continue;
    }
    lastWasPreference = false;
    // Days with nothing but connectors after them share the wording of the day that
    // closes the run ("בשני רביעי וחמישי"); wording before the first day of the line
    // applies to the run too ("לא יכול בראשון, בשני יכול החל משעה 12:00").
    const hasContent = t => /\d|לא|אין|יכול|פנוי|אפשרי|זמינ|עדיפות|נוח|מעדיפ|מלמד/.test(t);
    const runs = [];
    mentions.forEach((m, i) => {
      const own = segment.slice(m.end, i + 1 < mentions.length ? mentions[i + 1].index : segment.length);
      const prev = runs[runs.length - 1];
      if (prev && !prev.closed) prev.items.push({ days: m.days, own });
      else runs.push({ prefix: i === 0 ? segment.slice(0, m.index) : '', items: [{ days: m.days, own }] });
      runs[runs.length - 1].closed = hasContent(own);
    });

    const ruleFor = (q, ownText, prefix) => {
      const windows = parseWindows(ownText);
      if (/לא\s*(?:יכול|פנוי|זמינ|מתאים)|אין/.test(q)) return 'no';
      if (windows.length) return windows;
      if (/עדיפות|הכי נוח|מעדיפ|נוח ל/.test(q)) return 'pref';
      if (/מלמד|עובד|לימודים/.test(prefix)) return 'no';
      return [['00:00', '24:00']];
    };
    const assign = (days, rule) => {
      result.recognized = true;
      if (rule === 'pref') { result.preferredDays.push(...days); lastWasPreference = true; return; }
      if (rule === 'no') hasNegative = true; else hasListing = true;
      for (const d of days) result.days[d] = rule;
    };
    for (const run of runs) {
      const last = run.items[run.items.length - 1];
      const before = run.items.slice(0, -1).flatMap(it => it.days);
      if (before.length && hasContent(run.prefix) && hasContent(last.own)) {
        assign(before, ruleFor(run.prefix, '', run.prefix));
        assign(last.days, ruleFor(last.own, last.own, ''));
      } else {
        assign(run.items.flatMap(it => it.days), ruleFor(`${run.prefix} ${last.own}`, last.own, run.prefix));
      }
    }
  }

  // Only free times were listed ("פנויה ב: ראשון אחרי 12...") - the other weekdays are taken.
  if (hasListing && !hasNegative) {
    for (let d = 0; d <= 4; d++) if (!(d in result.days)) result.days[d] = 'no';
  }
  result.preferredDays = [...new Set(result.preferredDays)].sort();
  return result;
}

function detectRoles(guideCell, notes) {
  const n = String(notes || '');
  const magic = /קוס(?:ם|מת)(?!\s*בפוטנציה)/.test(n);
  const lecture = /הרצא|מרצה/.test(n);
  const onlyMagic = /רק\s+קוס/.test(n);
  let lab;
  if (yes(guideCell)) lab = true;
  else if (no(guideCell)) lab = false;
  else lab = !onlyMagic && !(lecture && !magic);
  return { lab: lab && !onlyMagic, magic, lecture: lecture && !onlyMagic };
}

function detectTags(notes) {
  const n = String(notes || '');
  const tags = [];
  if (/רק\s+קורסים/.test(n)) tags.push('רק קורסים');
  if (/ערבית/.test(n)) tags.push('ערבית');
  if (/חדש/.test(n)) tags.push('חדש/ה');
  if (/קוס(?:ם|מת)\s*בפוטנציה/.test(n)) tags.push('קוסם/ת בפוטנציה');
  if (/לא\s+הגיע/.test(n)) tags.push('לא הגיע/ה להכשרות');
  if (/מילואים/.test(n)) tags.push('מילואים');
  return tags;
}

/**
 * @param rows  sheet rows as arrays of cell values (first row with "שם" is the header)
 * @returns instructors
 */
export function parseRosterRows(rows, today = new Date()) {
  const headerIndex = rows.findIndex(r => r.some(c => /שם/.test(String(c || ''))));
  if (headerIndex < 0) throw new Error('לא נמצאה שורת כותרת עם עמודת "שם"');
  const header = rows[headerIndex].map(c => String(c || '').trim());
  const col = {};
  for (const [key, re] of Object.entries(COLUMNS)) {
    const i = header.findIndex(h => re.test(h));
    if (i >= 0) col[key] = i;
  }
  const get = (row, key) => (col[key] === undefined ? '' : String(row[col[key]] ?? '').trim());

  const out = [];
  for (const row of rows.slice(headerIndex + 1)) {
    const name = get(row, 'name').replace(/\?+$/, '').replace(/\s+/g, ' ').trim();
    if (!name) continue;
    const notes = get(row, 'notes');
    const safety = get(row, 'safety');
    const safetyOk = yes(safety) || /מאשר/.test(safety);
    out.push({
      id: normalizeName(name),
      name,
      phone: get(row, 'phone'),
      degree: get(row, 'degree'),
      notes,
      safetyOk,
      safetyNote: safetyOk || !safety ? '' : safety,
      roles: detectRoles(get(row, 'guide'), notes),
      active: !/כבר לא|לא רוצה להדריך|לא פעיל|עזב/.test(notes),
      tags: detectTags(notes),
      availability: parseAvailability(notes, today),
    });
  }
  return out;
}

/** Day of week (0 = Sunday) of an ISO date, independent of the local time zone. */
export function weekday(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d)).getUTCDay();
}

/**
 * Can this instructor take a slot?
 * @returns {{level: 'ok'|'soft'|'partial'|'no', reason: string}}
 */
export function checkAvailability(instructor, date, start, end) {
  const a = instructor.availability || {};
  if (instructor.active === false) return { level: 'no', reason: 'לא פעיל/ה' };
  if (a.unavailableUntil && date <= a.unavailableUntil) {
    const [y, m, d] = a.unavailableUntil.split('-');
    return { level: 'no', reason: `לא זמין/ה עד ${d}/${m}/${y.slice(2)}` };
  }
  const day = weekday(date);
  const rule = a.days ? a.days[day] : undefined;
  const s = toMinutes(start), e = toMinutes(end);
  if (rule === 'no') return { level: 'no', reason: `לא פנוי/ה ביום ${DAY_NAMES[day]}` };
  if (Array.isArray(rule)) {
    const ws = rule.map(([f, t]) => [toMinutes(f), toMinutes(t)]);
    const text = rule.map(([f, t]) => (f === '00:00' ? `עד ${t}` : t === '24:00' ? `מ-${f}` : `${f}-${t}`)).join(', ');
    if (!ws.some(([f, t]) => f <= s && t >= e)) {
      const overlap = ws.some(([f, t]) => f < e && t > s);
      return { level: overlap ? 'partial' : 'no', reason: `ביום ${DAY_NAMES[day]} פנוי/ה ${text}` };
    }
  }
  if (a.preferredDays && a.preferredDays.length && !a.preferredDays.includes(day)) {
    return { level: 'soft', reason: `מעדיפ/ה ימי ${a.preferredDays.map(d => DAY_NAMES[d]).join(', ')}` };
  }
  if (a.preferredHours) {
    const [f, t] = a.preferredHours.map(toMinutes);
    if (s < f || e > t) return { level: 'soft', reason: `מעדיפ/ה שעות ${a.preferredHours[0]}-${a.preferredHours[1]}` };
  }
  return { level: 'ok', reason: '' };
}

/** Short Hebrew summary of the weekly rules, for the instructors table. */
export function describeAvailability(a) {
  if (!a) return '';
  const parts = [];
  for (let d = 0; d <= 5; d++) {
    const rule = a.days?.[d];
    if (rule === 'no') parts.push(`${DAY_NAMES[d]}: לא`);
    else if (Array.isArray(rule) && !(rule.length === 1 && rule[0][0] === '00:00' && rule[0][1] === '24:00')) {
      parts.push(`${DAY_NAMES[d]}: ${rule.map(([f, t]) => (f === '00:00' ? `עד ${t}` : t === '24:00' ? `מ-${f}` : `${f}-${t}`)).join(', ')}`);
    }
  }
  if (a.preferredDays?.length) parts.push(`עדיפות: ${a.preferredDays.map(d => DAY_NAMES[d]).join(', ')}`);
  if (a.preferredHours) parts.push(`שעות מועדפות: ${a.preferredHours[0]}-${a.preferredHours[1]}`);
  if (a.unavailableUntil) parts.push(`לא זמין/ה עד ${a.unavailableUntil.split('-').reverse().join('/')}`);
  return parts.join(' · ');
}
