// Scheduling logic that does not touch the page: merging a re-imported schedule
// with the assignments already made, which slots are shown, conflicts, workload,
// ranking candidates for a position and filling open positions automatically.

import { normalizeName } from './names.js';
import { toMinutes } from './schedule-parser.js';
import { checkAvailability } from './roster.js';

export const ROLE_OF_SLOT = { lab: 'lab', lecture: 'lecture', magic: 'magic' };

/**
 * Positions in the app carry who filled them (the schedule file, the user, or
 * auto-fill) and what the schedule file said there at the last import, so the
 * next import can tell the coordinator's changes from the user's own choices.
 */
export function toAppEvent(parsed) {
  return {
    ...parsed,
    ignored: false,
    note: '',
    changes: [],
    slots: parsed.slots.map(s => ({
      ...s,
      groups: s.groups.map(g => ({ ...g, positions: g.positions.map(name => ({ name, source: name ? 'pdf' : null, pdfName: name })) })),
    })),
  };
}

// Words every title shares; they say nothing about which school or course it is.
const GENERIC_WORDS = new Set([
  'יום', 'עיון', 'שיא', 'בכימיה', 'כימיה', 'בביולוגיה', 'ביולוגיה', 'בביוטכנולוגיה', 'לתיכון', 'תיכון', 'לחטיבת', 'חטיבת',
  'ביניים', 'כיתה', 'כיתות', 'תל', 'תלמידים', 'תלמידות', 'שתי', 'קבוצות', 'קבוצה', 'סהכ', 'שא', 'במשך', 'מפגש', 'קורס',
  'שכבה', 'שכבות', 'הגיעו', 'לאולפנת', 'אולפנת', 'מתחלפות',
].map(normalizeName));

function titleWords(title, all = false) {
  return new Set(normalizeName(title).split(' ').filter(w => all || ((w.length > 1 || /\d/.test(w)) && !GENERIC_WORDS.has(w))));
}

/**
 * Share of the shorter title's distinctive words found in the other: a title that
 * only gained details still matches. A title with no distinctive words at all
 * ("יום עיון בכימיה") is compared word for word and must match almost entirely.
 */
function titleScore(a, b) {
  let A = titleWords(a), B = titleWords(b);
  let scale = 1;
  if (!A.size || !B.size) { A = titleWords(a, true); B = titleWords(b, true); scale = 0.6 / 0.9; }
  if (!A.size || !B.size) return { score: 0, common: 0 };
  let common = 0;
  for (const w of A) if (B.has(w)) common++;
  return { score: (common / Math.min(A.size, B.size)) * scale, common: scale === 1 ? common : 0 };
}

const sessionNumber = title => (String(title).match(/מפגש\s*(\d+)/) || [])[1] || '';
const sameName = (a, b) => normalizeName(a) === normalizeName(b);
const ddmm = iso => `${iso.slice(8, 10)}/${iso.slice(5, 7)}`;
const namesIn = slot => slot.groups.flatMap(g => g.positions).filter(p => p.name).map(p => p.name);
const daysBetween = (a, b) => Math.abs(Date.parse(a) - Date.parse(b)) / 86400000;

/** Pair each fresh event with the event it was in the previous import, if any. */
export function pairEvents(oldEvents, fresh, from, to) {
  const pairs = new Map();
  const usedOld = new Set();
  const link = (ev, old, moved = false) => { pairs.set(ev, { old, moved }); usedOld.add(old); };

  for (const ev of fresh) {
    const old = oldEvents.find(o => o.id === ev.id && !usedOld.has(o));
    if (old) link(ev, old);
  }
  // Same day, title changed a little: best matches first.
  const candidates = [];
  for (const ev of fresh) {
    if (pairs.has(ev)) continue;
    for (const o of oldEvents) {
      if (usedOld.has(o) || o.manual || o.date !== ev.date) continue;
      const { score } = titleScore(o.title, ev.title);
      if (score >= 0.6 && sessionNumber(o.title) === sessionNumber(ev.title)) candidates.push({ score, ev, o });
    }
  }
  candidates.sort((a, b) => b.score - a.score);
  for (const { ev, o } of candidates) if (!pairs.has(ev) && !usedOld.has(o)) link(ev, o);

  // Moved to another day: an event that vanished from its date and one clear match that appeared elsewhere.
  const lostOld = oldEvents.filter(o => !usedOld.has(o) && !o.manual && o.date >= from && o.date <= to && o.slots.length);
  const looksSame = (o, ev) => {
    const { score, common } = titleScore(o.title, ev.title);
    return score >= 0.8 && common >= 2 && sessionNumber(o.title) === sessionNumber(ev.title) && daysBetween(o.date, ev.date) <= 45;
  };
  for (const o of lostOld) {
    const options = fresh.filter(ev => !pairs.has(ev) && ev.date !== o.date && looksSame(o, ev));
    if (options.length !== 1) continue;
    const rivals = lostOld.filter(x => x !== o && !usedOld.has(x) && looksSame(x, options[0]));
    if (!rivals.length) link(options[0], o, true);
  }
  return { pairs, usedOld };
}

/** Pair each fresh slot with its earlier version: same times, then same activity, then same place among its kind. */
export function pairSlots(ev, old) {
  const pairs = [];
  const used = new Set();
  const paired = s => pairs.some(p => p[0] === s);
  const pass = test => {
    for (const s of ev.slots) {
      if (paired(s)) continue;
      const o = old.slots.find(x => !used.has(x) && x.kind === s.kind && test(s, x));
      if (o) { used.add(o); pairs.push([s, o]); }
    }
  };
  pass((s, o) => o.start === s.start && o.end === s.end && o.label === s.label);
  pass((s, o) => o.start === s.start && o.end === s.end);
  pass((s, o) => o.label === s.label);
  for (const kind of new Set(ev.slots.map(s => s.kind))) {
    const restNew = ev.slots.filter(s => s.kind === kind && !paired(s));
    const restOld = old.slots.filter(o => o.kind === kind && !used.has(o) && !o.manual);
    if (restNew.length && restNew.length === restOld.length) restNew.forEach((s, i) => { used.add(restOld[i]); pairs.push([s, restOld[i]]); });
  }
  return { pairs, used };
}

/**
 * One position: what the file says now (p), what the app had (op).
 * The user's choice stands while the file keeps saying what it said before;
 * when the coordinator writes a different name, the file wins and it is reported.
 */
function mergePosition(p, op, slot, changes) {
  const fileNow = p.name;
  if (!op) return;
  const mine = (op.source === 'user' || op.source === 'auto') && op.name;
  const fileBefore = op.pdfName !== undefined ? op.pdfName : op.source === 'pdf' ? op.name : fileNow;
  const where = `"${slot.label}" ${slot.start}-${slot.end}`;
  if (mine) {
    if (fileNow === fileBefore || !fileNow || (fileBefore && sameName(fileNow, fileBefore))) Object.assign(p, { name: op.name, source: op.source, at: op.at });
    else if (sameName(fileNow, op.name)) p.source = 'pdf';
    else changes.push({ level: 'attention', text: `בלו"ז רשום/ה ${fileNow} במקום ${op.name} ששובץ/ה אצלך (${where})` });
    return;
  }
  if (op.name && !fileNow) changes.push({ level: 'attention', text: `${op.name} הוסר/ה מהלו"ז - העמדה פתוחה (${where})` });
  else if (op.name && fileNow && !sameName(op.name, fileNow)) changes.push({ level: 'info', text: `בלו"ז: ${fileNow} במקום ${op.name} (${where})` });
  else if (!op.name && fileNow) changes.push({ level: 'info', text: `נוסף/ה בלו"ז: ${fileNow} (${where})` });
}

function mergeEvent(ev, old, moved) {
  const changes = [];
  ev.ignored = old.ignored;
  ev.note = old.note;
  if (old.at) ev.at = old.at;
  if (old.kindOverride) { ev.kind = old.kindOverride; ev.kindOverride = old.kindOverride; }
  if (moved) {
    const people = old.slots.flatMap(namesIn);
    changes.push({ level: 'attention', text: `הוזז מ-${ddmm(old.date)} ל-${ddmm(ev.date)}${people.length ? ` - לבדוק שהמשובצים פנויים בתאריך החדש (${[...new Set(people)].join(', ')})` : ''}` });
  }
  const { pairs, used } = pairSlots(ev, old);
  for (const [slot, oldSlot] of pairs) {
    // The user's own edits to the shape of a slot win over the file.
    if (oldSlot.edited) {
      Object.assign(slot, { label: oldSlot.label, location: oldSlot.location, kind: oldSlot.kind, start: oldSlot.start, end: oldSlot.end, subject: oldSlot.subject, edited: true, at: oldSlot.at });
      slot.groups = oldSlot.groups.map(g => ({ ...g, positions: g.positions.map(p => ({ ...p })) }));
      continue;
    }
    if ((slot.start !== oldSlot.start || slot.end !== oldSlot.end) && namesIn(oldSlot).length) {
      changes.push({ level: 'attention', text: `השעות של "${slot.label}" השתנו: ${oldSlot.start}-${oldSlot.end} ← ${slot.start}-${slot.end} (משובצים: ${namesIn(oldSlot).join(', ')})` });
    }
    slot.groups.forEach((g, gi) => {
      const og = oldSlot.groups[gi];
      g.positions.forEach((p, pi) => mergePosition(p, og && og.positions[pi], slot, changes));
      // Positions the user added beyond the file's.
      if (og) for (let pi = g.positions.length; pi < og.positions.length; pi++) g.positions.push({ ...og.positions[pi] });
    });
    for (let gi = slot.groups.length; gi < oldSlot.groups.length; gi++) {
      slot.groups.push({ ...oldSlot.groups[gi], positions: oldSlot.groups[gi].positions.map(p => ({ ...p })) });
    }
  }
  for (const o of old.slots) {
    if (used.has(o)) continue;
    if (o.manual) { ev.slots.push(o); continue; }
    const people = namesIn(o);
    if (people.length) changes.push({ level: 'attention', text: `"${o.label}" ${o.start}-${o.end} ירד מהלו"ז (היו משובצים: ${people.join(', ')})` });
  }
  ev.slots.sort((a, b) => toMinutes(a.start) - toMinutes(b.start));
  ev.changes = changes;
}

/**
 * Merge a freshly parsed schedule into the current events.
 * Choices the user made (assignments, hidden events, notes) survive; changes the
 * coordinator made (times, names, dates) are applied and listed in the report.
 * Old events in the date range of the new file that are gone from it are dropped,
 * unless someone was assigned there - those stay, marked as missing.
 * @returns {{events, added, updated, removed, kept, report: {evId, date, title, level, text}[]}}
 */
export function mergeSchedule(oldEvents, newParsed) {
  const fresh = newParsed.map(toAppEvent);
  if (!fresh.length) return { events: oldEvents, added: 0, updated: 0, removed: 0, kept: [], report: [] };
  const dates = fresh.map(e => e.date).sort();
  const from = dates[0], to = dates[dates.length - 1];
  const { pairs, usedOld } = pairEvents(oldEvents, fresh, from, to);
  let added = 0, updated = 0;

  for (const ev of fresh) {
    const pair = pairs.get(ev);
    if (!pair) {
      added++;
      // Nothing to report on the first import.
      ev.changes = oldEvents.length ? [{ level: 'new', text: 'חדש בלו"ז' }] : [];
      continue;
    }
    updated++;
    mergeEvent(ev, pair.old, pair.moved);
  }

  const kept = [];
  let removed = 0;
  for (const o of oldEvents) {
    if (usedOld.has(o)) continue;
    const inRange = o.date >= from && o.date <= to;
    const people = o.slots.flatMap(namesIn);
    if (o.manual || !inRange) kept.push({ ...o, changes: o.missingFromSchedule ? o.changes : [] });
    else if (people.length && o.kind !== 'other' && !o.ignored) {
      kept.push({
        ...o,
        missingFromSchedule: true,
        changes: [{ level: 'attention', text: `לא מופיע בלו"ז העדכני (בוטל או הוזז?) - היו משובצים: ${[...new Set(people)].join(', ')}` }],
      });
    } else removed++;
  }
  const events = [...kept, ...fresh].sort((a, b) => a.date.localeCompare(b.date) || firstStart(a) - firstStart(b));
  const report = events.flatMap(ev => (ev.changes || []).map(c => ({ evId: ev.id, date: ev.date, title: ev.title, ...c })));
  return { events, added, updated, removed, kept: kept.filter(e => e.missingFromSchedule), report };
}

const firstStart = e => (e.slots.length ? Math.min(...e.slots.map(s => toMinutes(s.start))) : 24 * 60);

/** Is the event (and which of its slots) shown under the current filters? */
export function visibleSlots(ev, settings) {
  return ev.slots.filter(s => settings.slotKinds[s.kind] !== false && settings.subjects[s.subject || 'general'] !== false);
}

export function isRelevant(ev) {
  return ev.kind !== 'other' && ev.slots.length > 0;
}

export function eventMatches(ev, settings) {
  if (settings.month && settings.month !== 'all' && !ev.date.startsWith(settings.month)) return false;
  if (settings.showHidden) return true;
  if (ev.ignored || !isRelevant(ev)) return false;
  if (settings.kinds[ev.kind] === false) return false;
  const slots = visibleSlots(ev, settings);
  if (!slots.length) return false;
  if (settings.onlyOpen && !slots.some(s => s.groups.some(g => g.positions.some(p => !p.name)))) return false;
  return true;
}

/** Every filled position, flattened, for load and conflict checks. */
export function allAssignments(events) {
  const out = [];
  for (const ev of events) {
    if (ev.ignored) continue;
    for (const s of ev.slots) {
      s.groups.forEach((g, gi) => g.positions.forEach((p, pi) => {
        if (p.name) out.push({ name: p.name, key: normalizeName(p.name), ev, slot: s, gi, pi, start: toMinutes(s.start), end: toMinutes(s.end) });
      }));
    }
  }
  return out;
}

/** Assignments of a person that overlap a slot in time (other than the position itself). */
export function conflictsFor(assignments, name, ev, slot, gi, pi) {
  const key = normalizeName(name);
  const s = toMinutes(slot.start), e = toMinutes(slot.end);
  return assignments.filter(a => a.key === key && a.ev.date === ev.date && a.start < e && a.end > s
    && !(a.slot === slot && a.gi === gi && a.pi === pi));
}

export function loadByPerson(assignments, month) {
  const load = new Map();
  for (const a of assignments) {
    if (month && month !== 'all' && !a.ev.date.startsWith(month)) continue;
    load.set(a.key, (load.get(a.key) || 0) + 1);
  }
  return load;
}

/** How often two people worked in the same group - pairs that work well together. */
export function pairCounts(events) {
  const pairs = new Map();
  for (const ev of events) for (const s of ev.slots) for (const g of s.groups) {
    const names = g.positions.map(p => p.name).filter(Boolean).map(normalizeName);
    for (const a of names) for (const b of names) if (a !== b) pairs.set(`${a}|${b}`, (pairs.get(`${a}|${b}`) || 0) + 1);
  }
  return pairs;
}

const LEVEL_ORDER = { ok: 0, soft: 1, partial: 2, no: 3 };

/**
 * Rank the instructors for one position.
 * @returns {{instructor, tier: 'good'|'maybe'|'busy'|'role', reasons: string[], load: number, partner: number, availability}[]}
 */
export function rankCandidates({ instructors, events, ev, slot, gi, pi, month, assignments, pairs }) {
  assignments = assignments || allAssignments(events);
  pairs = pairs || pairCounts(events);
  const load = loadByPerson(assignments, month || ev.date.slice(0, 7));
  const groupMates = slot.groups[gi].positions.filter((p, i) => i !== pi && p.name).map(p => normalizeName(p.name));
  const role = ROLE_OF_SLOT[slot.kind];

  return instructors.map(instructor => {
    const key = normalizeName(instructor.name);
    const reasons = [];
    let tier = 'good';
    const availability = checkAvailability(instructor, ev.date, slot.start, slot.end);
    const clash = conflictsFor(assignments, instructor.name, ev, slot, gi, pi);
    const sameGroup = groupMates.includes(key);

    if (!instructor.roles?.[role]) {
      tier = 'role';
      reasons.push(role === 'magic' ? 'לא קוסם/ת' : role === 'lecture' ? 'לא מרצה' : 'לא מסומן/ת כמדריך/ה');
    }
    if (sameGroup) { tier = 'busy'; reasons.push('כבר בקבוצה הזו'); }
    else if (clash.length) {
      tier = tier === 'role' ? 'role' : 'busy';
      reasons.push(`משובץ/ת במקביל: ${clash[0].slot.start}-${clash[0].slot.end} ${clash[0].slot.label}`);
    }
    if (availability.level === 'no') { if (tier !== 'role') tier = 'busy'; reasons.push(availability.reason); }
    else if (availability.level !== 'ok') { if (tier === 'good') tier = 'maybe'; reasons.push(availability.reason); }
    if (role === 'lab' && !instructor.safetyOk) { if (tier === 'good') tier = 'maybe'; reasons.push('הדרכת בטיחות לא הושלמה'); }
    if (instructor.tags?.includes('רק קורסים') && ev.kind !== 'course') { if (tier === 'good') tier = 'maybe'; reasons.push('רוצה רק קורסים'); }
    if (instructor.tags?.includes('לא הגיע/ה להכשרות')) { if (tier === 'good') tier = 'maybe'; reasons.push('לא הגיע/ה להכשרות'); }
    const partner = groupMates.reduce((n, m) => n + (pairs.get(`${key}|${m}`) || 0), 0);
    return { instructor, tier, reasons: [...new Set(reasons)], load: load.get(key) || 0, partner, availability };
  }).sort((a, b) => {
    const t = ['good', 'maybe', 'busy', 'role'];
    // Fewest assignments first; a usual partner in the group counts like one less.
    const score = c => c.load * 2 - Math.min(c.partner, 3);
    return t.indexOf(a.tier) - t.indexOf(b.tier)
      || score(a) - score(b)
      || LEVEL_ORDER[a.availability.level] - LEVEL_ORDER[b.availability.level]
      || a.instructor.name.localeCompare(b.instructor.name, 'he');
  });
}

/**
 * Fill every open position in the given events with the best available candidate
 * (role fits, free at that time, fewest assignments this month, preferred partners).
 * @returns {number} how many positions were filled
 */
export function autoFill({ instructors, events, targets, settings }) {
  let filled = 0;
  const pairs = pairCounts(events);
  const order = [...targets].sort((a, b) => a.date.localeCompare(b.date) || firstStart(a) - firstStart(b));
  for (const ev of order) {
    for (const slot of visibleSlots(ev, settings)) {
      slot.groups.forEach((g, gi) => g.positions.forEach((p, pi) => {
        if (p.name) return;
        const ranked = rankCandidates({ instructors, events, ev, slot, gi, pi, month: ev.date.slice(0, 7), pairs });
        const pick = ranked.find(c => c.tier === 'good') || ranked.find(c => c.tier === 'maybe' && c.availability.level === 'soft');
        if (!pick) return;
        p.name = pick.instructor.name;
        p.source = 'auto';
        filled++;
      }));
    }
  }
  return filled;
}

export function openCount(ev, settings) {
  let open = 0, total = 0;
  for (const s of visibleSlots(ev, settings)) for (const g of s.groups) for (const p of g.positions) { total++; if (!p.name) open++; }
  return { open, total };
}
