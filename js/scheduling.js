// Scheduling logic that does not touch the page: merging a re-imported schedule
// with the assignments already made, which slots are shown, conflicts, workload,
// ranking candidates for a position and filling open positions automatically.

import { normalizeName } from './names.js';
import { toMinutes } from './schedule-parser.js';
import { checkAvailability } from './roster.js';

export const ROLE_OF_SLOT = { lab: 'lab', lecture: 'lecture', magic: 'magic' };

/** Positions in the app carry who filled them: the schedule file, the user, or auto-fill. */
export function toAppEvent(parsed) {
  return {
    ...parsed,
    ignored: false,
    note: '',
    slots: parsed.slots.map(s => ({
      ...s,
      groups: s.groups.map(g => ({ ...g, positions: g.positions.map(name => ({ name, source: name ? 'pdf' : null })) })),
    })),
  };
}

const words = s => new Set(normalizeName(s).split(' ').filter(w => w.length > 1));
function similarity(a, b) {
  const A = words(a), B = words(b);
  if (!A.size || !B.size) return 0;
  let common = 0;
  for (const w of A) if (B.has(w)) common++;
  return common / Math.min(A.size, B.size);
}

function findOldEvent(ev, oldEvents, taken) {
  const byId = oldEvents.find(o => o.id === ev.id && !taken.has(o));
  if (byId) return byId;
  let best = null, bestScore = 0.6;
  for (const o of oldEvents) {
    if (taken.has(o) || o.date !== ev.date || o.manual) continue;
    const score = similarity(o.title, ev.title);
    if (score > bestScore) { best = o; bestScore = score; }
  }
  return best;
}

function findOldSlot(slot, index, oldSlots, taken) {
  const same = oldSlots.filter(o => !taken.has(o) && o.kind === slot.kind);
  return same.find(o => o.start === slot.start && o.end === slot.end && o.label === slot.label)
    || same.find(o => o.start === slot.start && o.end === slot.end)
    || same.find(o => o.id === slot.id)
    || null;
}

/**
 * Merge a freshly parsed schedule into the current events.
 * Choices the user made (assignments, hidden events, notes) survive; old events
 * in the date range of the new file that are gone from it are dropped, unless the
 * user assigned someone there - those stay, marked as missing from the schedule.
 */
export function mergeSchedule(oldEvents, newParsed) {
  const fresh = newParsed.map(toAppEvent);
  if (!fresh.length) return { events: oldEvents, added: 0, updated: 0, removed: 0, kept: [] };
  const dates = fresh.map(e => e.date).sort();
  const from = dates[0], to = dates[dates.length - 1];
  const taken = new Set();
  let added = 0, updated = 0;

  for (const ev of fresh) {
    const old = findOldEvent(ev, oldEvents, taken);
    if (!old) { added++; continue; }
    taken.add(old);
    updated++;
    ev.ignored = old.ignored;
    ev.note = old.note;
    if (old.kindOverride) { ev.kind = old.kindOverride; ev.kindOverride = old.kindOverride; }
    const usedSlots = new Set();
    ev.slots.forEach((slot, i) => {
      const oldSlot = findOldSlot(slot, i, old.slots, usedSlots);
      if (!oldSlot) return;
      usedSlots.add(oldSlot);
      // The user's own edits to the shape of a slot win over the file.
      if (oldSlot.edited) {
        Object.assign(slot, { label: oldSlot.label, location: oldSlot.location, kind: oldSlot.kind, start: oldSlot.start, end: oldSlot.end, edited: true });
        slot.groups = oldSlot.groups.map(g => ({ ...g, positions: g.positions.map(p => ({ ...p })) }));
        return;
      }
      slot.groups.forEach((g, gi) => {
        const og = oldSlot.groups[gi];
        if (!og) return;
        g.positions.forEach((p, pi) => {
          const op = og.positions[pi];
          if (op && op.name && (op.source === 'user' || op.source === 'auto' || !p.name)) Object.assign(p, op);
        });
        // Positions the user added beyond the file's.
        for (let pi = g.positions.length; pi < og.positions.length; pi++) g.positions.push({ ...og.positions[pi] });
      });
      for (let gi = slot.groups.length; gi < oldSlot.groups.length; gi++) {
        slot.groups.push({ ...oldSlot.groups[gi], positions: oldSlot.groups[gi].positions.map(p => ({ ...p })) });
      }
    });
    // Slots the user added by hand.
    for (const s of old.slots) if (s.manual && !usedSlots.has(s)) ev.slots.push(s);
  }

  const kept = [];
  let removed = 0;
  for (const o of oldEvents) {
    if (taken.has(o)) continue;
    const inRange = o.date >= from && o.date <= to;
    const hasUserWork = o.slots.some(s => s.groups.some(g => g.positions.some(p => p.name && p.source !== 'pdf')));
    if (o.manual || !inRange) kept.push(o);
    else if (hasUserWork) kept.push({ ...o, missingFromSchedule: true });
    else removed++;
  }
  const events = [...kept, ...fresh].sort((a, b) => a.date.localeCompare(b.date) || firstStart(a) - firstStart(b));
  return { events, added, updated, removed, kept: kept.filter(e => e.missingFromSchedule) };
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
