// Keeping one person's work in step across devices without a server of our own.
// Every change the user makes is stamped with the time it was made, and two
// copies of the data merge position by position: the latest choice wins,
// wherever it was made, and nothing done on either device is thrown away.

import { pairEvents, pairSlots } from './scheduling.js';
import { toMinutes } from './schedule-parser.js';

/** This device's view preferences: kept per device, never sent. */
export const LOCAL_FIELDS = ['settings', 'exportOptions'];
export const FILE_APP = 'nsm-shibutz';

const TOMBSTONE_DAYS = 120;
const clone = o => JSON.parse(JSON.stringify(o));
const stamp = x => (x && x.at) || 0;
const time = iso => Date.parse(iso || '') || 0;

const eventFields = e => JSON.stringify([!!e.ignored, e.note || '', e.kindOverride || null]);
const slotShape = s => JSON.stringify([s.start, s.end, s.kind, s.label, s.location || '', s.subject || null, s.groups.map(g => [g.label || '', g.positions.length])]);
const posValue = p => JSON.stringify([p.name || null, p.name ? p.source || null : null]);
const withoutStamp = x => JSON.stringify({ ...x, at: undefined });

/**
 * Stamp what changed between two versions of the state with `now`, and carry
 * the earlier stamps over to objects that were rebuilt without them.
 * @returns {number} how many things changed
 */
export function stampChanges(prev, next, now = Date.now()) {
  let changed = 0;
  const mark = (x, isChanged, before) => {
    if (isChanged) { x.at = now; changed++; } else if (!x.at && before && before.at) x.at = before.at;
  };
  const prevEvents = new Map((prev?.events || []).map(e => [e.id, e]));
  for (const ev of next.events) {
    const pe = prevEvents.get(ev.id);
    if (pe) mark(ev, eventFields(ev) !== eventFields(pe), pe);
    else if (ev.manual && !ev.at) mark(ev, true);
    const prevSlots = new Map((pe?.slots || []).map(s => [s.id, s]));
    for (const s of ev.slots) {
      const ps = prevSlots.get(s.id);
      if (ps) mark(s, slotShape(s) !== slotShape(ps), ps);
      else if ((s.manual || s.edited) && !s.at) mark(s, true);
      s.groups.forEach((g, gi) => g.positions.forEach((p, pi) => {
        const pp = ps && ps.groups[gi] && ps.groups[gi].positions[pi];
        if (pp) mark(p, posValue(p) !== posValue(pp), pp);
        else if (p.name && !p.at) mark(p, true);
      }));
    }
  }
  const prevEdits = prev?.instructorEdits || {};
  for (const [id, e] of Object.entries(next.instructorEdits || {})) {
    const pe = prevEdits[id];
    mark(e, !pe || withoutStamp(e) !== withoutStamp(pe), pe);
  }
  const prevExtra = new Map((prev?.extraInstructors || []).map(i => [i.id, i]));
  for (const i of next.extraInstructors || []) {
    const pi = prevExtra.get(i.id);
    mark(i, !pi || withoutStamp(i) !== withoutStamp(pi), pi);
  }
  if (next.notes) {
    const text = n => JSON.stringify([n?.gaps || '', n?.tasks || '']);
    mark(next.notes, text(next.notes) !== text(prev?.notes), prev?.notes);
  }
  return changed;
}

/** The part of the state that travels between devices. */
export function sharedState(state) {
  const out = { ...state };
  for (const k of LOCAL_FIELDS) delete out[k];
  return out;
}

/** A file or sync payload: who wrote it, when, and the shared state. */
export function packState(state, device) {
  return { app: FILE_APP, format: 1, savedAt: new Date().toISOString(), device, state: sharedState(state) };
}

/** The state inside a payload, or an old-style backup (the bare state). */
export function unpackState(data) {
  const s = data && data.app === FILE_APP ? data.state : data;
  if (!s || !Array.isArray(s.events) || !Array.isArray(s.instructors)) return null;
  return s;
}

const scheduleImports = s => (s.imports && s.imports.schedule) || [];
const lastImportAt = s => Math.max(0, ...scheduleImports(s).map(i => time(i.at)));
const rosterAt = s => time(s.imports && s.imports.roster && s.imports.roster.at);
const namesIn = slot => slot.groups.flatMap(g => g.positions).filter(p => p.name).map(p => p.name);
const latestSlotStamp = s => Math.max(stamp(s), ...s.groups.flatMap(g => g.positions.map(stamp)));
const latestEventStamp = ev => Math.max(stamp(ev), ...ev.slots.map(latestSlotStamp));
const firstStart = e => (e.slots.length ? Math.min(...e.slots.map(s => toMinutes(s.start))) : 24 * 60);

/** Copy into `target` every position the other side changed more recently. */
function takeNewerPositions(target, source) {
  target.forEach((g, gi) => g.positions.forEach((p, pi) => {
    const sp = source[gi] && source[gi].positions[pi];
    if (sp && stamp(sp) > stamp(p)) Object.assign(p, { name: sp.name ?? null, source: sp.source ?? null, at: sp.at });
  }));
}

function mergeSlot(slot, other) {
  if (stamp(other) > stamp(slot)) {
    // The other side reshaped this slot more recently: its shape, then the newest name per position.
    const mine = slot.groups;
    Object.assign(slot, clone(other), { id: slot.id });
    takeNewerPositions(slot.groups, mine);
  } else {
    takeNewerPositions(slot.groups, other.groups);
  }
}

function mergeEvent(ev, other) {
  if (stamp(other) > stamp(ev)) {
    ev.ignored = other.ignored;
    ev.note = other.note;
    if (other.kindOverride) { ev.kindOverride = other.kindOverride; ev.kind = other.kindOverride; }
    ev.at = other.at;
  }
  // Same slot id first (manual slots keep theirs), then the import's own matching.
  const pairs = [];
  const used = new Set();
  for (const s of ev.slots) {
    const o = other.slots.find(x => x.id === s.id && !used.has(x));
    if (o) { pairs.push([s, o]); used.add(o); }
  }
  const rest = pairSlots({ slots: ev.slots.filter(s => !pairs.some(p => p[0] === s)) }, { slots: other.slots.filter(o => !used.has(o)) });
  for (const [s, o] of rest.pairs) { pairs.push([s, o]); used.add(o); }
  for (const [s, o] of pairs) mergeSlot(s, o);
  for (const o of other.slots) if (!used.has(o) && o.manual) ev.slots.push(clone(o));
  ev.slots.sort((a, b) => toMinutes(a.start) - toMinutes(b.start));
}

function mergeDeleted(a = {}, b = {}, now) {
  const out = { events: {}, slots: {} };
  const cutoff = now - TOMBSTONE_DAYS * 86400000;
  for (const kind of ['events', 'slots']) {
    for (const src of [a[kind] || {}, b[kind] || {}]) {
      for (const [id, at] of Object.entries(src)) if (at > cutoff && at > (out[kind][id] || 0)) out[kind][id] = at;
    }
  }
  return out;
}

function newerById(a = [], b = []) {
  const map = new Map(a.map(x => [x.id, x]));
  for (const x of b) if (!map.has(x.id) || stamp(x) > stamp(map.get(x.id))) map.set(x.id, x);
  return [...map.values()].map(clone);
}

/**
 * Merge the state from another device (or a file) into this device's state.
 * The side with the newer schedule import gives the events their shape; on top
 * of it every position, event note and instructor edit takes the most recent
 * value from either side. Deletions are kept as dated tombstones.
 */
export function mergeStates(local, remote, now = Date.now()) {
  // On a tie the remote side is the base, so two devices settle on the same
  // copy instead of handing an unstamped difference back and forth.
  const remoteNewer = lastImportAt(remote) >= lastImportAt(local);
  const base = clone(remoteNewer ? remote : local);
  const other = remoteNewer ? local : remote;
  const baseImportAt = lastImportAt(base);

  // Events: pair, merge, then add what only the other side has.
  const dates = [...base.events, ...other.events].map(e => e.date).sort();
  const { pairs } = pairEvents(other.events, base.events, dates[0] || '', dates[dates.length - 1] || '');
  const usedOther = new Set();
  for (const ev of base.events) {
    const pair = pairs.get(ev);
    if (!pair) continue;
    usedOther.add(pair.old);
    mergeEvent(ev, pair.old);
  }
  const latest = scheduleImports(base).reduce((a, b) => (time(b.at) > time(a?.at) ? b : a), null);
  for (const o of other.events) {
    if (usedOther.has(o)) continue;
    const inLatest = latest && latest.from && o.date >= latest.from && o.date <= latest.to;
    if (o.manual || !inLatest) { base.events.push(clone(o)); continue; }
    // The newer schedule no longer has it: keep it only if people were assigned there.
    const people = o.slots.flatMap(namesIn);
    if (people.length && o.kind !== 'other' && !o.ignored) {
      base.events.push({
        ...clone(o),
        missingFromSchedule: true,
        changes: [{ level: 'attention', text: `לא מופיע בלו"ז העדכני (בוטל או הוזז?) - היו משובצים: ${[...new Set(people)].join(', ')}` }],
      });
    }
  }

  // Deletions: a manual or missing item stays deleted; an item the schedule
  // file still lists comes back once a newer file is imported.
  const deleted = mergeDeleted(local.deleted, remote.deleted, now);
  const gone = (at, item, stampOf) => at && at > stampOf(item) && (item.manual || item.missingFromSchedule || at > baseImportAt);
  base.events = base.events.filter(ev => !gone(deleted.events[ev.id], ev, latestEventStamp));
  for (const ev of base.events) ev.slots = ev.slots.filter(s => !gone(deleted.slots[s.id], s, latestSlotStamp));
  base.events.sort((a, b) => a.date.localeCompare(b.date) || firstStart(a) - firstStart(b));
  base.deleted = deleted;

  // Instructors: the newer roster import, then edits and added people one by one.
  if (rosterAt(other) > rosterAt(base)) {
    base.instructors = clone(other.instructors);
    base.imports = { ...base.imports, roster: clone(other.imports.roster) };
  }
  base.extraInstructors = newerById(base.extraInstructors, other.extraInstructors);
  const edits = { ...(base.instructorEdits || {}) };
  for (const [id, e] of Object.entries(other.instructorEdits || {})) if (!edits[id] || stamp(e) > stamp(edits[id])) edits[id] = clone(e);
  base.instructorEdits = edits;
  if (stamp(other.notes) > stamp(base.notes)) base.notes = clone(other.notes);

  // Import history and the last import's report.
  const seen = new Set();
  base.imports = {
    ...base.imports,
    schedule: [...scheduleImports(local), ...scheduleImports(remote)]
      .filter(i => { const k = `${i.name}|${i.at}`; if (seen.has(k)) return false; seen.add(k); return true; })
      .sort((a, b) => time(a.at) - time(b.at))
      .slice(-10)
      .map(clone),
  };
  if (time(other.lastImport?.at) > time(base.lastImport?.at)) base.lastImport = clone(other.lastImport);

  for (const k of LOCAL_FIELDS) if (local[k] !== undefined) base[k] = clone(local[k]);
  base.updatedAt = Math.max(local.updatedAt || 0, remote.updatedAt || 0);
  return base;
}

/** What a merge changed on this device, for a one-line summary. */
export function describeChanges(before, after) {
  const positions = s => {
    const m = new Map();
    for (const ev of s.events) for (const sl of ev.slots) {
      sl.groups.forEach((g, gi) => g.positions.forEach((p, pi) => m.set(`${ev.date}|${sl.start}|${sl.end}|${sl.label}|${gi}|${pi}`, p.name || '')));
    }
    return m;
  };
  const A = positions(before), B = positions(after);
  let assigned = 0, cleared = 0;
  for (const [k, v] of B) if (v && (A.get(k) || '') !== v) assigned++;
  for (const [k, v] of A) if (v && B.has(k) && !B.get(k)) cleared++;
  const evKey = e => `${e.date}|${e.title}`;
  const had = new Set(before.events.map(evKey));
  const events = after.events.filter(e => !had.has(evKey(e))).length;
  const editsBefore = before.instructorEdits || {};
  const people = Object.entries(after.instructorEdits || {}).filter(([id, e]) => withoutStamp(e) !== withoutStamp(editsBefore[id] || {})).length
    + (after.extraInstructors || []).filter(i => !(before.extraInstructors || []).some(x => x.id === i.id)).length;
  return {
    assigned, cleared, events, people,
    roster: rosterAt(after) > rosterAt(before),
    schedule: lastImportAt(after) > lastImportAt(before),
    notes: JSON.stringify([after.notes?.gaps, after.notes?.tasks]) !== JSON.stringify([before.notes?.gaps, before.notes?.tasks]),
    same: JSON.stringify(sharedState(before)) === JSON.stringify(sharedState(after)),
  };
}
