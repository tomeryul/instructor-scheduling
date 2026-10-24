// The page: import the two files, assign instructors slot by slot, keep track
// of instructors' details, and export the "משימות"-style document.

import { extractScheduleCells } from './pdf-grid.js';
import { parseCells, EVENT_KINDS, SLOT_KINDS, SUBJECTS } from './schedule-parser.js';
import { createNameMatcher, normalizeName } from './names.js';
import { parseRosterRows, parseWindows, checkAvailability, describeAvailability, DAY_NAMES, weekday } from './roster.js';
import {
  mergeSchedule, eventMatches, visibleSlots, isRelevant, allAssignments, conflictsFor,
  rankCandidates, autoFill, openCount, ROLE_OF_SLOT,
} from './scheduling.js';
import { buildDocument, toHtml, toPlainText, toDocxBlob, dayLabel, shortDate, monthLabel } from './export.js';
import { openStore, saveFile, loadScript } from './storage.js';
import { demoState } from './demo.js';

const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const DOCX_URL = 'https://cdn.jsdelivr.net/npm/docx@9.0.2/build/index.umd.js';

const DEFAULT_SETTINGS = {
  subjects: { chem: true, general: true, bio: false, cs: false, other: false },
  kinds: { study: true, course: true },
  slotKinds: { lab: true, lecture: true, magic: true },
  month: 'all',
  onlyOpen: false,
  showHidden: false,
  showPast: false,
};
const DEFAULT_EXPORT = { month: 'all', location: false, sections: { assigned: true, open: true, pool: true, courses: true, done: false } };
const ROLE_LABELS = { lab: 'מדריך/ה', magic: 'קוסם/ת', lecture: 'מרצה' };

const $ = sel => document.querySelector(sel);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const clone = o => JSON.parse(JSON.stringify(o));

let store = null;
let state = emptyState();
let demo = false;
let view = 'board';
let filtersOpen = window.matchMedia('(min-width: 700px)').matches;
let autoUndo = null;
const notices = { roster: null, schedule: null };

function emptyState() {
  return {
    version: 1,
    instructors: [],
    extraInstructors: [],
    instructorEdits: {},
    events: [],
    settings: clone(DEFAULT_SETTINGS),
    notes: { gaps: '', tasks: '' },
    exportOptions: clone(DEFAULT_EXPORT),
    imports: { roster: null, schedule: [] },
  };
}

function withDefaults(s) {
  const base = emptyState();
  return {
    ...base,
    ...s,
    settings: { ...base.settings, ...(s.settings || {}), subjects: { ...base.settings.subjects, ...(s.settings?.subjects || {}) }, kinds: { ...base.settings.kinds, ...(s.settings?.kinds || {}) }, slotKinds: { ...base.settings.slotKinds, ...(s.settings?.slotKinds || {}) } },
    exportOptions: { ...base.exportOptions, ...(s.exportOptions || {}), sections: { ...base.exportOptions.sections, ...(s.exportOptions?.sections || {}) } },
    notes: { ...base.notes, ...(s.notes || {}) },
    imports: { ...base.imports, ...(s.imports || {}) },
  };
}

const hasData = s => s.events.length > 0 || s.instructors.length > 0 || s.extraInstructors.length > 0;

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// ---------------------------------------------------------------- people

/** Instructors as imported, with the user's edits on top. */
function people() {
  return [...state.instructors, ...state.extraInstructors].map(i => {
    const e = state.instructorEdits[i.id];
    if (!e) return i;
    return { ...i, ...e, roles: { ...i.roles, ...(e.roles || {}) }, availability: e.availability || i.availability };
  });
}

let personIndex = null;
function personByName(name) {
  if (!personIndex) {
    const list = people();
    personIndex = { list, byKey: new Map(list.map(p => [normalizeName(p.name), p])), matcher: createNameMatcher(list.map(p => p.name)) };
  }
  const direct = personIndex.byKey.get(normalizeName(name));
  if (direct) return direct;
  const m = personIndex.matcher.match(name);
  return m ? personIndex.byKey.get(normalizeName(m)) : null;
}

function editPerson(id, patch) {
  state.instructorEdits[id] = { ...(state.instructorEdits[id] || {}), ...patch };
}

// ---------------------------------------------------------------- saving

let saveTimer = null;
let saving = false;
let dirty = false;

function commit() {
  personIndex = null;
  render();
  if (demo) return;
  dirty = true;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flush, 500);
}

async function flush() {
  if (saving || !dirty || !store) return;
  saving = true;
  dirty = false;
  try {
    await store.save(state);
  } catch (e) {
    toast('השמירה נכשלה: ' + (e.message || e.code || e));
  } finally {
    saving = false;
    if (dirty) flush();
  }
}

function leaveDemo() {
  if (!demo) return;
  demo = false;
  state = emptyState();
  autoUndo = null;
  $('#demo-banner').hidden = true;
}

// ---------------------------------------------------------------- feedback

let toastTimer = null;
function toast(msg) {
  const el = $('#toast');
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, 3200);
}

async function copyText(text, html) {
  try {
    if (html && window.ClipboardItem && navigator.clipboard?.write) {
      await navigator.clipboard.write([new ClipboardItem({
        'text/plain': new Blob([text], { type: 'text/plain' }),
        'text/html': new Blob([`<div dir="rtl">${html}</div>`], { type: 'text/html' }),
      })]);
    } else {
      await navigator.clipboard.writeText(text);
    }
    toast('הועתק');
  } catch {
    openDialog(`
      <div class="sheet-head"><h2 id="dialog-title">העתקה</h2><p class="muted small">ההעתקה האוטומטית לא זמינה כאן. הטקסט מסומן - העתיקו אותו (Ctrl+C).</p></div>
      <div class="sheet-body"><textarea class="input" id="copy-area" rows="14" readonly>${esc(text)}</textarea></div>
      <div class="sheet-foot"><button class="btn" data-close>סגירה</button></div>`);
    const ta = $('#copy-area');
    ta.focus();
    ta.select();
  }
}

// ---------------------------------------------------------------- dialog

const dialog = () => $('#dialog');
let dialogHandlers = {};

function openDialog(html, handlers = {}) {
  const d = dialog();
  d.innerHTML = `<div class="sheet-inner">${html}</div>`;
  dialogHandlers = handlers;
  if (!d.open) d.showModal();
  const focus = d.querySelector('[autofocus]');
  if (focus) focus.focus();
}

function closeDialog() {
  const d = dialog();
  if (d.open) d.close();
  dialogHandlers = {};
}

// ---------------------------------------------------------------- lookups

function findEvent(id) { return state.events.find(e => e.id === id); }
function findSlot(ev, slotId) { return ev.slots.find(s => s.id === slotId); }

function monthsInData() {
  return [...new Set(state.events.filter(isRelevant).map(e => e.date.slice(0, 7)))].sort();
}

function eventShown(ev) {
  const s = state.settings;
  if (!eventMatches(ev, s)) return false;
  if (!s.showPast && ev.date < todayISO()) return false;
  return true;
}

function slotsShown(ev) {
  return state.settings.showHidden && (ev.ignored || !isRelevant(ev)) ? ev.slots : visibleSlots(ev, state.settings);
}

/** Problems with a filled position: wrong role, not available, double-booked. */
function positionIssues(ev, slot, gi, pi, name, assignments) {
  const person = personByName(name);
  if (!person) return { level: 'external', reasons: ['לא ברשימת המדריכים'] };
  const reasons = [];
  let level = 'ok';
  const role = ROLE_OF_SLOT[slot.kind];
  if (!person.roles?.[role]) { reasons.push(`לא מסומן/ת כ${ROLE_LABELS[role]}`); level = 'warn'; }
  const av = checkAvailability(person, ev.date, slot.start, slot.end);
  if (av.level === 'no') { reasons.push(av.reason); level = 'bad'; }
  else if (av.level !== 'ok') { reasons.push(av.reason); if (level === 'ok') level = 'warn'; }
  const clash = conflictsFor(assignments, name, ev, slot, gi, pi);
  if (clash.length) { reasons.push(`משובץ/ת במקביל ב-${clash[0].slot.start}-${clash[0].slot.end} (${clash[0].ev === ev ? 'באותו אירוע' : clash[0].ev.title.slice(0, 40)})`); level = 'bad'; }
  if (role === 'lab' && !person.safetyOk) { reasons.push('הדרכת בטיחות לא הושלמה'); if (level === 'ok') level = 'warn'; }
  return { level, reasons };
}

// ---------------------------------------------------------------- board view

function renderBoard() {
  const el = $('#view-board');
  if (!hasData(state)) {
    el.innerHTML = `
      <div class="panel empty-state">
        <h2>מתחילים מהקבצים</h2>
        <p>ייבאו את רשימת המדריכים (Excel) ואת קובץ הלו"ז (PDF). האתר יאתר את ימי העיון, הקורסים, ההרצאות ומופעי הקסמים, ותוכלו לשבץ מדריכים בלחיצה.</p>
        <button class="btn btn-primary" data-go="files">לייבוא הקבצים</button>
      </div>`;
    return;
  }
  const s = state.settings;
  const today = todayISO();
  const months = monthsInData();
  const assignments = allAssignments(state.events);
  const shown = state.events.filter(eventShown);

  // Counts for the filter chips, over the month in view.
  const inMonth = state.events.filter(ev => isRelevant(ev) && !ev.ignored && (s.month === 'all' || ev.date.startsWith(s.month)) && (s.showPast || ev.date >= today));
  const subjectCounts = {};
  const slotKindCounts = {};
  const kindCounts = {};
  for (const ev of inMonth) {
    kindCounts[ev.kind] = (kindCounts[ev.kind] || 0) + 1;
    for (const sl of ev.slots) {
      const subj = sl.subject || 'general';
      subjectCounts[subj] = (subjectCounts[subj] || 0) + 1;
      slotKindCounts[sl.kind] = (slotKindCounts[sl.kind] || 0) + 1;
    }
  }
  const pastCount = state.events.filter(ev => isRelevant(ev) && !ev.ignored && ev.date < today && (s.month === 'all' || ev.date.startsWith(s.month))).length;
  const hiddenCount = state.events.filter(ev => (ev.ignored || !isRelevant(ev)) && (s.month === 'all' || ev.date.startsWith(s.month))).length;

  let open = 0, total = 0;
  for (const ev of shown) { const c = openCount(ev, s); open += c.open; total += c.total; }
  const pct = total ? Math.round(((total - open) / total) * 100) : 0;

  const chip = (group, key, label, count, swatch) => `
    <label class="chip"><input type="checkbox" data-filter="${group}" value="${key}" ${s[group][key] !== false ? 'checked' : ''}>
      ${swatch ? `<span class="swatch" style="background:${swatch}"></span>` : ''}${esc(label)}${count !== undefined ? ` <span class="count">${count}</span>` : ''}</label>`;
  const toggle = (key, label, count) => `
    <label class="chip"><input type="checkbox" data-toggle="${key}" ${s[key] ? 'checked' : ''}>${esc(label)}${count !== undefined ? ` <span class="count">${count}</span>` : ''}</label>`;

  const days = [];
  for (const ev of shown) {
    const last = days[days.length - 1];
    if (last && last.date === ev.date) last.events.push(ev);
    else days.push({ date: ev.date, events: [ev] });
  }

  el.innerHTML = `
    <div class="panel board-bar">
      <div class="summary">
        <div class="stat open"><b>${open}</b><span>עמדות פתוחות</span></div>
        <div class="stat done"><b>${total - open}</b><span>משובצות</span></div>
        <div class="stat"><b>${shown.length}</b><span>אירועים</span></div>
        <div class="progress" title="${pct}% משובץ"><span style="width:${pct}%"></span></div>
        <div class="actions">
          ${autoUndo ? '<button class="btn" data-act="undo-auto">ביטול השיבוץ האוטומטי</button>' : ''}
          <button class="btn btn-primary" data-act="auto" ${open ? '' : 'disabled'}>שיבוץ אוטומטי של החסרים</button>
          <button class="btn" data-act="add-event">אירוע ידני</button>
        </div>
      </div>
      <details class="filters-box" ${filtersOpen ? 'open' : ''}>
      <summary>סינון: ${s.month === 'all' ? 'כל החודשים' : monthLabel(s.month, true)} · ${Object.entries(SUBJECTS).filter(([k]) => s.subjects[k] !== false).map(([, v]) => v).join(', ')}${s.onlyOpen ? ' · רק עם חסרים' : ''}</summary>
      <div class="filters">
        <div class="filter-group"><span class="label">חודש</span>
          <select class="input" id="f-month" data-setting="month">
            <option value="all">כל החודשים</option>
            ${months.map(m => `<option value="${m}" ${s.month === m ? 'selected' : ''}>${monthLabel(m, true)}</option>`).join('')}
          </select>
        </div>
        <div class="filter-group"><span class="label">סוג</span>
          ${chip('kinds', 'study', 'ימי עיון', kindCounts.study || 0, 'var(--study)')}
          ${chip('kinds', 'course', 'קורסים', kindCounts.course || 0, 'var(--course)')}
        </div>
        <div class="filter-group"><span class="label">משבצות</span>
          ${chip('slotKinds', 'lab', 'הדרכות', slotKindCounts.lab || 0, 'var(--study)')}
          ${chip('slotKinds', 'lecture', 'הרצאות', slotKindCounts.lecture || 0, 'var(--lecture)')}
          ${chip('slotKinds', 'magic', 'מופעי קסמים', slotKindCounts.magic || 0, 'var(--magic)')}
        </div>
        <div class="filter-group"><span class="label">תחום</span>
          ${Object.entries(SUBJECTS).map(([k, label]) => chip('subjects', k, label, subjectCounts[k] || 0)).join('')}
        </div>
        <div class="filter-group">
          ${toggle('onlyOpen', 'רק עם חסרים')}
          ${toggle('showPast', 'כולל אירועים שעברו', pastCount)}
          ${toggle('showHidden', 'הצג מוסתרים ולא רלוונטיים', hiddenCount)}
        </div>
      </div>
      </details>
      <div class="legend">
        <span><i style="background:var(--bad-soft);border:1px dashed var(--bad)"></i>חסר</span>
        <span><i style="background:var(--ok-soft)"></i>משובץ</span>
        <span><i style="background:var(--accent-soft)"></i>שובץ אוטומטית</span>
        <span><i style="background:var(--warn-soft);border-color:var(--warn)"></i>יש הסתייגות</span>
        <span><i style="background:var(--bad-soft);border-color:var(--bad)"></i>לא זמין/ה או כפילות</span>
      </div>
    </div>
    ${days.length ? days.map(d => `
      <section class="day">
        <div class="day-head"><h2>יום ${DAY_NAMES[weekday(d.date)]}</h2><span class="date">${shortDate(d.date)}</span>${d.date < today ? '<span class="past">עבר</span>' : ''}</div>
        ${d.events.map(ev => renderEvent(ev, assignments)).join('')}
      </section>`).join('') : `
      <div class="panel empty-state"><h2>אין אירועים להצגה</h2><p>אין אירועים שמתאימים לסינון הנוכחי. אפשר לשנות את החודש או את התחומים למעלה.</p></div>`}`;
}

function renderEvent(ev, assignments) {
  const c = openCount(ev, state.settings);
  const kind = ev.kind || 'other';
  const slots = slotsShown(ev);
  const subjects = [...new Set(slots.map(sl => sl.subject).filter(Boolean))].map(k => SUBJECTS[k]);
  const hidden = ev.ignored || !isRelevant(ev);
  return `
  <article class="ev kind-${kind} ${hidden ? 'is-ignored' : ''}" data-ev="${esc(ev.id)}">
    <div class="ev-head">
      <div>
        <div class="ev-tags">
          <span class="tag tag-kind">${EVENT_KINDS[kind]}</span>
          ${subjects.map(t => `<span class="tag">${esc(t)}</span>`).join('')}
          ${hidden ? '<span class="tag">מוסתר</span>' : c.open ? `<span class="tag tag-open">חסרים ${c.open}</span>` : c.total ? '<span class="tag tag-done">שובץ</span>' : ''}
          ${ev.missingFromSchedule ? '<span class="tag tag-warn" title="האירוע לא נמצא בקובץ הלו&quot;ז האחרון שיובא">לא מופיע בלו"ז העדכני</span>' : ''}
          ${ev.manual ? '<span class="tag">ידני</span>' : ''}
        </div>
        <h3 class="ev-title clamp" data-act="expand" title="לחצו להצגת הכותרת המלאה">${esc(ev.title || '(ללא כותרת)')}</h3>
      </div>
      <div class="ev-actions">
        <button class="btn btn-ghost btn-small" data-act="note">${ev.note ? 'עריכת הערה' : 'הערה'}</button>
        ${ev.raw ? '<button class="btn btn-ghost btn-small" data-act="raw">טקסט מהלו"ז</button>' : ''}
        <button class="btn btn-ghost btn-small" data-act="add-slot">משבצת</button>
        ${hidden ? `<button class="btn btn-ghost btn-small" data-act="unhide">${ev.ignored ? 'החזרה לשיבוץ' : 'סימון כרלוונטי'}</button>` : '<button class="btn btn-ghost btn-small" data-act="hide">הסתרה</button>'}
      </div>
    </div>
    ${slots.length ? `<ol class="slots">${slots.map(sl => renderSlot(ev, sl, assignments)).join('')}</ol>` : '<p class="muted small" style="margin:0 14px 12px">אין משבצות שדורשות מדריכים. אפשר להוסיף משבצת ידנית.</p>'}
    ${ev.note ? `<p class="ev-note">${esc(ev.note)}</p>` : ''}
    <pre class="raw" hidden>${esc(ev.raw || '')}</pre>
  </article>`;
}

function renderSlot(ev, sl, assignments) {
  return `
  <li class="slot slot-${sl.kind}" data-slot="${esc(sl.id)}">
    <span class="time">${esc(sl.start)}–${esc(sl.end)}</span>
    <div class="slot-main">
      <div><span class="slot-kind"><span class="swatch"></span>${SLOT_KINDS[sl.kind]}</span><span class="slot-label">${esc(sl.label)}</span>
        ${sl.location ? `<span class="loc"> · ${esc(sl.location)}</span>` : ''}</div>
      <div class="groups">
        ${sl.groups.map((g, gi) => `
          <div class="group">
            ${g.label ? `<span class="g-label">${esc(g.label)}${g.activity && g.activity !== sl.label ? ' · ' + esc(g.activity) : ''}</span>` : ''}
            ${g.positions.map((p, pi) => renderPosition(ev, sl, gi, pi, p, assignments)).join('')}
          </div>`).join('')}
      </div>
    </div>
    <div class="slot-tools"><button class="icon-btn" data-act="edit-slot" title="עריכת המשבצת">עריכה</button></div>
  </li>`;
}

function renderPosition(ev, sl, gi, pi, p, assignments) {
  const ref = `data-pos="${esc(sl.id)}|${gi}|${pi}"`;
  if (!p.name) return `<button class="pos empty" ${ref} title="לחצו לבחירת ${sl.kind === 'lecture' ? 'מרצה' : 'מדריך/ה'}">＋ ${sl.kind === 'lecture' ? 'מרצה' : sl.kind === 'magic' ? 'קוסם/ת' : 'מדריך/ה'}</button>`;
  const issues = positionIssues(ev, sl, gi, pi, p.name, assignments);
  const cls = issues.level === 'bad' ? 'bad' : issues.level === 'warn' ? 'warn' : issues.level === 'external' ? 'filled external' : p.source === 'auto' ? 'filled src-auto' : 'filled';
  const title = [p.source === 'auto' ? 'שובץ אוטומטית' : p.source === 'pdf' ? 'מופיע בלו"ז' : 'שובץ ידנית', ...issues.reasons].join(' · ');
  return `<button class="pos ${cls}" ${ref} title="${esc(title)}">${issues.level === 'ok' || issues.level === 'external' ? '' : '<span class="dot"></span>'}${esc(p.name)}</button>`;
}

// ---------------------------------------------------------------- picker

function openPicker(ev, slot, gi, pi) {
  const assignments = allAssignments(state.events);
  const ranked = rankCandidates({ instructors: people(), events: state.events, ev, slot, gi, pi, month: ev.date.slice(0, 7), assignments });
  const current = slot.groups[gi].positions[pi].name;
  const group = slot.groups[gi];
  const tiers = [
    ['good', 'מתאימים ופנויים'],
    ['maybe', 'אפשריים, עם הסתייגות'],
    ['busy', 'לא זמינים או משובצים במקביל'],
    ['role', 'בלי התפקיד המתאים'],
  ];
  const row = c => `
    <button class="cand tier-${c.tier}" data-pick="${esc(c.instructor.name)}" data-search="${esc(normalizeName(c.instructor.name))}">
      <span class="name">${esc(c.instructor.name)}${c.instructor.name === current ? ' (משובץ/ת עכשיו)' : ''}</span>
      <span class="meta">${c.load} החודש</span>
      ${c.reasons.length || c.partner ? `<span class="reasons">${c.partner ? '<span class="good">עובדים יחד הרבה</span>' : ''}${c.reasons.map(r => `<span>${esc(r)}</span>`).join('')}</span>` : ''}
    </button>`;
  const sections = tiers.map(([tier, title], i) => {
    const list = ranked.filter(c => c.tier === tier);
    if (!list.length) return '';
    const body = `<div class="cand-list">${list.map(row).join('')}</div>`;
    return i < 2
      ? `<div class="cand-section"><h3>${title} (${list.length})</h3>${body}</div>`
      : `<details class="more cand-section"><summary>${title} (${list.length})</summary>${body}</details>`;
  }).join('');
  const roleName = slot.kind === 'lecture' ? 'מרצה' : slot.kind === 'magic' ? 'קוסם/ת' : 'מדריך/ה';

  openDialog(`
    <div class="sheet-head">
      <h2 id="dialog-title">בחירת ${roleName}</h2>
      <p class="small muted">${esc(dayLabel(ev.date))} · ${esc(slot.start)}–${esc(slot.end)} · ${esc(slot.label)}${group.label ? ' · ' + esc(group.label) : ''}</p>
      <p class="small muted">${esc(ev.title.slice(0, 120))}</p>
      <input class="input" id="pick-search" type="search" placeholder="חיפוש שם…" autocomplete="off" autofocus>
    </div>
    <div class="sheet-body" id="pick-body">${sections || '<p class="muted">אין מדריכים ברשימה. ייבאו את קובץ המדריכים או כתבו שם למטה.</p>'}</div>
    <div class="sheet-foot">
      <input class="input grow" id="pick-free" placeholder="שם שלא ברשימה" autocomplete="off">
      <button class="btn" data-act="pick-free">שיבוץ שם זה</button>
      ${current ? '<button class="btn btn-danger" data-act="pick-clear">ניקוי השיבוץ</button>' : ''}
      <button class="btn btn-ghost" data-close>ביטול</button>
    </div>`, {
    pick(name) {
      const p = slot.groups[gi].positions[pi];
      p.name = name;
      p.source = 'user';
      closeDialog();
      commit();
    },
    clear() {
      const p = slot.groups[gi].positions[pi];
      p.name = null;
      p.source = null;
      closeDialog();
      commit();
    },
  });
}

// ---------------------------------------------------------------- slot & event editors

function slotEditor(ev, slot, isNew) {
  const draft = clone(slot);
  const renderGroups = () => draft.groups.map((g, gi) => `
    <div class="group-edit" data-gi="${gi}">
      <label class="field">שם קבוצה<input class="input" data-g="label" value="${esc(g.label)}" placeholder="למשל קבוצה א'"></label>
      <span class="small">עמדות: <b class="num">${g.positions.length}</b></span>
      <button class="btn btn-small" data-gact="minus" ${g.positions.length <= 1 ? 'disabled' : ''}>−</button>
      <button class="btn btn-small" data-gact="plus">+</button>
      ${draft.groups.length > 1 ? '<button class="btn btn-small btn-danger" data-gact="remove">הסרת קבוצה</button>' : ''}
    </div>`).join('');
  const html = () => `
    <div class="sheet-head"><h2 id="dialog-title">${isNew ? 'משבצת חדשה' : 'עריכת משבצת'}</h2><p class="small muted">${esc(dayLabel(ev.date))} · ${esc(ev.title.slice(0, 90))}</p></div>
    <div class="sheet-body">
      <div class="form-grid">
        <label class="field">התחלה<input class="input" type="time" id="se-start" value="${esc(draft.start)}"></label>
        <label class="field">סיום<input class="input" type="time" id="se-end" value="${esc(draft.end)}"></label>
        <label class="field">סוג<select class="input" id="se-kind">${Object.entries(SLOT_KINDS).map(([k, v]) => `<option value="${k}" ${draft.kind === k ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
        <label class="field">תחום<select class="input" id="se-subject">${Object.entries(SUBJECTS).map(([k, v]) => `<option value="${k}" ${(draft.subject || 'general') === k ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
      </div>
      <label class="field">תיאור הפעילות<input class="input" id="se-label" value="${esc(draft.label)}"></label>
      <label class="field">מיקום<input class="input" id="se-location" value="${esc(draft.location || '')}"></label>
      <div id="se-groups" style="display:grid;gap:8px">${renderGroups()}</div>
      <div><button class="btn btn-small" data-gact="add-group">הוספת קבוצה</button></div>
    </div>
    <div class="sheet-foot">
      <button class="btn btn-primary" data-act="slot-save">שמירה</button>
      ${isNew ? '' : '<button class="btn btn-danger" data-act="slot-delete">מחיקת המשבצת</button>'}
      <button class="btn btn-ghost" data-close>ביטול</button>
    </div>`;
  const readForm = () => {
    draft.start = $('#se-start').value || draft.start;
    draft.end = $('#se-end').value || draft.end;
    draft.kind = $('#se-kind').value;
    draft.subject = $('#se-subject').value === 'general' ? null : $('#se-subject').value;
    draft.label = $('#se-label').value.trim() || SLOT_KINDS[draft.kind];
    draft.location = $('#se-location').value.trim();
    document.querySelectorAll('#se-groups [data-gi]').forEach(row => { draft.groups[+row.dataset.gi].label = row.querySelector('[data-g="label"]').value.trim(); });
  };
  openDialog(html(), {
    group(action, gi) {
      readForm();
      if (action === 'add-group') draft.groups.push({ label: '', activity: '', positions: [{ name: null, source: null }, { name: null, source: null }] });
      else if (action === 'plus') draft.groups[gi].positions.push({ name: null, source: null });
      else if (action === 'minus') draft.groups[gi].positions.pop();
      else if (action === 'remove') draft.groups.splice(gi, 1);
      $('#se-groups').innerHTML = renderGroups();
    },
    save() {
      readForm();
      draft.edited = true;
      if (isNew) { draft.manual = true; ev.slots.push(draft); }
      else Object.assign(slot, draft);
      ev.slots.sort((a, b) => a.start.localeCompare(b.start));
      closeDialog();
      commit();
    },
    remove() {
      ev.slots = ev.slots.filter(s => s !== slot);
      closeDialog();
      commit();
    },
  });
}

function newSlot(ev) {
  return {
    id: `${ev.id}.m${Date.now().toString(36)}`, start: '10:00', end: '11:30', kind: 'lab', label: 'הדרכה', location: '',
    subject: ev.subject || null, groups: [{ label: '', activity: '', positions: [{ name: null, source: null }, { name: null, source: null }] }],
  };
}

function eventEditor() {
  openDialog(`
    <div class="sheet-head"><h2 id="dialog-title">אירוע ידני</h2><p class="small muted">לאירוע שלא מופיע בקובץ הלו"ז. אחרי השמירה מוסיפים לו משבצות.</p></div>
    <div class="sheet-body">
      <div class="form-grid">
        <label class="field">תאריך<input class="input" type="date" id="ee-date" value="${todayISO()}"></label>
        <label class="field">סוג<select class="input" id="ee-kind"><option value="study">יום עיון</option><option value="course">קורס</option></select></label>
      </div>
      <label class="field">כותרת<input class="input" id="ee-title" placeholder="למשל: יום עיון בכימיה לתיכון… (איש קשר) כיתה י״א"></label>
    </div>
    <div class="sheet-foot"><button class="btn btn-primary" data-act="event-save">יצירה והוספת משבצת</button><button class="btn btn-ghost" data-close>ביטול</button></div>`, {
    save() {
      const date = $('#ee-date').value;
      const title = $('#ee-title').value.trim();
      if (!date || !title) { toast('צריך תאריך וכותרת'); return; }
      leaveDemo();
      const ev = { id: `m${Date.now().toString(36)}`, date, color: null, kind: $('#ee-kind').value, subject: null, title, raw: '', slots: [], ignored: false, note: '', manual: true };
      state.events.push(ev);
      state.events.sort((a, b) => a.date.localeCompare(b.date));
      closeDialog();
      commit();
      slotEditor(ev, newSlot(ev), true);
    },
  });
}

function noteEditor(ev) {
  openDialog(`
    <div class="sheet-head"><h2 id="dialog-title">הערה לאירוע</h2><p class="small muted">${esc(dayLabel(ev.date))} · ${esc(ev.title.slice(0, 90))}. ההערה תופיע גם בקובץ המיוצא.</p></div>
    <div class="sheet-body"><textarea class="input" id="note-text" rows="4" placeholder="למשל: תוזכרו ביום חמישי / השעות הוקדמו, כל המדריכים עודכנו" autofocus>${esc(ev.note || '')}</textarea></div>
    <div class="sheet-foot"><button class="btn btn-primary" data-act="note-save">שמירה</button><button class="btn btn-ghost" data-close>ביטול</button></div>`, {
    save() {
      ev.note = $('#note-text').value.trim();
      closeDialog();
      commit();
    },
  });
}

// ---------------------------------------------------------------- people view

let peopleQuery = '';

function windowsToText(rule) {
  if (rule === 'no') return 'לא';
  if (!Array.isArray(rule)) return '';
  return rule.map(([f, t]) => (f === '00:00' && t === '24:00' ? 'כל היום' : f === '00:00' ? `עד ${t}` : t === '24:00' ? `מ-${f}` : `${f}-${t}`)).join(', ');
}

function upcomingFor(person) {
  const today = todayISO();
  return allAssignments(state.events).filter(a => a.ev.date >= today && personByName(a.name) === person)
    .sort((a, b) => a.ev.date.localeCompare(b.ev.date) || a.start - b.start);
}

function renderPeople() {
  const el = $('#view-people');
  const list = people();
  if (!list.length) {
    el.innerHTML = `<div class="panel empty-state"><h2>אין עדיין מדריכים</h2><p>ייבאו את קובץ המדריכים (Excel) בלשונית "קבצים", או הוסיפו מדריך/ה ידנית.</p>
      <div style="display:flex;gap:8px;flex-wrap:wrap;justify-content:center"><button class="btn btn-primary" data-go="files">לייבוא</button><button class="btn" data-act="add-person">הוספת מדריך/ה</button></div></div>`;
    return;
  }
  const today = todayISO();
  const upcoming = new Map();
  for (const a of allAssignments(state.events)) {
    if (a.ev.date < today) continue;
    const p = personByName(a.name);
    if (p) upcoming.set(p.id, (upcoming.get(p.id) || 0) + 1);
  }
  const q = normalizeName(peopleQuery);
  const rows = list.filter(p => !q || normalizeName(p.name).includes(q));
  el.innerHTML = `
    <div class="panel" style="display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between">
      <div><h2>מדריכים</h2><p class="small muted" style="margin:0">התפקידים והזמינות נקראו מהעמודות "הדרכה" ו"הערות". אפשר לתקן כאן - התיקונים נשמרים גם כשמייבאים את הקובץ מחדש.</p></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <input class="input" type="search" id="people-search" placeholder="חיפוש…" value="${esc(peopleQuery)}">
        <button class="btn" data-act="add-person">הוספת מדריך/ה</button>
      </div>
    </div>
    <div class="table-scroll">
      <table class="people">
        <thead><tr><th>שם</th><th>טלפון</th><th>תפקידים</th><th>זמינות</th><th>בטיחות</th><th>שיבוצים קרובים</th><th>הערות מהקובץ</th><th></th></tr></thead>
        <tbody>${rows.map(p => `
          <tr class="${p.active === false ? 'inactive' : ''}" data-person="${esc(p.id)}">
            <td><b>${esc(p.name)}</b>${p.tags?.length ? `<div class="small muted">${p.tags.map(esc).join(' · ')}</div>` : ''}${p.degree ? `<div class="small muted">${esc(p.degree)}</div>` : ''}</td>
            <td class="phone">${esc(p.phone || '')}</td>
            <td><div class="role-toggles">${Object.entries(ROLE_LABELS).map(([k, v]) => `<label class="chip"><input type="checkbox" data-role="${k}" ${p.roles?.[k] ? 'checked' : ''}>${v}</label>`).join('')}
              <label class="chip"><input type="checkbox" data-active ${p.active !== false ? 'checked' : ''}>פעיל/ה</label></div></td>
            <td class="small">${esc(describeAvailability(p.availability)) || '<span class="muted">אין אילוצים</span>'}<div><button class="btn btn-ghost btn-small" data-act="edit-avail">עריכה</button></div></td>
            <td class="small">${p.safetyOk ? '<span class="ok">הושלמה</span>' : `<span class="warn">לא הושלמה</span>${p.safetyNote ? `<div class="muted">${esc(p.safetyNote.slice(0, 80))}</div>` : ''}`}
              <div><label class="chip" style="margin-top:4px"><input type="checkbox" data-safety ${p.safetyOk ? 'checked' : ''}>אושר</label></div></td>
            <td class="num">${upcoming.get(p.id) || 0}</td>
            <td class="notes-cell">${esc(p.notes || '')}</td>
            <td><button class="btn btn-small" data-act="remind">תזכורת</button></td>
          </tr>`).join('')}</tbody>
      </table>
    </div>`;
}

function availabilityEditor(person) {
  const a = person.availability || { days: {}, preferredDays: [] };
  openDialog(`
    <div class="sheet-head"><h2 id="dialog-title">זמינות - ${esc(person.name)}</h2>
      <p class="small muted">לכל יום: ריק = פנוי/ה כל היום · "לא" = לא פנוי/ה · שעות כמו "מ-12:00", "עד 14:00", "10:00-13:45" (אפשר כמה, מופרדים בפסיק).</p>
      ${person.notes ? `<p class="small" style="white-space:pre-wrap;background:var(--surface-2);padding:6px 8px;border-radius:6px;margin:0">${esc(person.notes)}</p>` : ''}</div>
    <div class="sheet-body">
      <div class="avail-grid">
        ${[0, 1, 2, 3, 4, 5].map(d => `
          <b>${DAY_NAMES[d]}</b>
          <input class="input" data-day="${d}" value="${esc(windowsToText(a.days?.[d]))}" placeholder="פנוי/ה">
          <label class="chip"><input type="checkbox" data-pref="${d}" ${a.preferredDays?.includes(d) ? 'checked' : ''}>מועדף</label>`).join('')}
      </div>
      <div class="form-grid">
        <label class="field">לא זמין/ה עד תאריך<input class="input" type="date" id="av-until" value="${esc(a.unavailableUntil || '')}"></label>
        <label class="field">שעות מועדפות<input class="input" id="av-hours" value="${a.preferredHours ? esc(a.preferredHours.join('-')) : ''}" placeholder="למשל 9:00-14:00"></label>
      </div>
    </div>
    <div class="sheet-foot"><button class="btn btn-primary" data-act="avail-save">שמירה</button>
      ${state.instructorEdits[person.id]?.availability ? '<button class="btn" data-act="avail-reset">חזרה למה שנקרא מהקובץ</button>' : ''}
      <button class="btn btn-ghost" data-close>ביטול</button></div>`, {
    save() {
      const days = {};
      document.querySelectorAll('#dialog [data-day]').forEach(inp => {
        const t = inp.value.trim();
        if (!t || /^כל היום$|^פנוי/.test(t)) return;
        if (/^לא/.test(t)) { days[inp.dataset.day] = 'no'; return; }
        const w = parseWindows(t.replace(/^מ(?=\d)/, 'מ-'));
        if (w.length) days[inp.dataset.day] = w;
      });
      const preferredDays = [...document.querySelectorAll('#dialog [data-pref]:checked')].map(i => +i.dataset.pref);
      const hours = parseWindows($('#av-hours').value.trim());
      editPerson(person.id, { availability: { days, preferredDays, preferredHours: hours[0] || null, unavailableUntil: $('#av-until').value || null, recognized: true } });
      closeDialog();
      commit();
    },
    reset() {
      const e = { ...state.instructorEdits[person.id] };
      delete e.availability;
      state.instructorEdits[person.id] = e;
      closeDialog();
      commit();
    },
  });
}

function personEditor() {
  openDialog(`
    <div class="sheet-head"><h2 id="dialog-title">מדריך/ה חדש/ה</h2></div>
    <div class="sheet-body">
      <div class="form-grid">
        <label class="field">שם<input class="input" id="pe-name" autofocus></label>
        <label class="field">טלפון<input class="input" id="pe-phone" inputmode="tel"></label>
      </div>
      <div class="role-toggles">${Object.entries(ROLE_LABELS).map(([k, v]) => `<label class="chip"><input type="checkbox" data-newrole="${k}" ${k === 'lab' ? 'checked' : ''}>${v}</label>`).join('')}</div>
    </div>
    <div class="sheet-foot"><button class="btn btn-primary" data-act="person-save">הוספה</button><button class="btn btn-ghost" data-close>ביטול</button></div>`, {
    save() {
      const name = $('#pe-name').value.trim();
      if (!name) { toast('צריך שם'); return; }
      leaveDemo();
      const roles = {};
      document.querySelectorAll('#dialog [data-newrole]').forEach(i => { roles[i.dataset.newrole] = i.checked; });
      state.extraInstructors.push({ id: normalizeName(name), name, phone: $('#pe-phone').value.trim(), degree: '', notes: '', safetyOk: true, safetyNote: '', roles, active: true, tags: ['נוסף ידנית'], availability: { days: {}, preferredDays: [], preferredHours: null, unavailableUntil: null } });
      closeDialog();
      commit();
    },
  });
}

function reminderText(person) {
  const items = upcomingFor(person);
  const first = person.name.split(' ')[0];
  if (!items.length) return `היי ${first}, כרגע אין לך שיבוצים קרובים.`;
  const lines = items.map(a => {
    const g = a.slot.groups[a.gi];
    return `• ${dayLabel(a.ev.date)}, ${a.slot.start}-${a.slot.end} - ${a.slot.label}${g.label ? ` (${g.label})` : ''}${a.slot.location ? `, ${a.slot.location}` : ''}\n  ${a.ev.title.split('(')[0].trim()}`;
  });
  return `היי ${first}, תזכורת לשיבוצים שלך בנוער שוחר מדע:\n${lines.join('\n')}\nתודה!`;
}

// ---------------------------------------------------------------- export view

function exportBlocks() {
  const o = state.exportOptions;
  return buildDocument({ events: state.events, instructors: people(), notes: state.notes, settings: state.settings },
    { month: o.month, today: todayISO(), location: o.location, sections: o.sections });
}

function exportTitle() {
  const m = state.exportOptions.month;
  return `שיבוץ מדריכים${m && m !== 'all' ? ' - ' + monthLabel(m, true) : ''}`;
}

function renderExport() {
  const el = $('#view-export');
  const o = state.exportOptions;
  const months = monthsInData();
  const sec = (key, label) => `<label><input type="checkbox" data-section="${key}" ${o.sections[key] !== false ? 'checked' : ''}> ${label}</label>`;
  el.innerHTML = `
    <div class="export-layout">
      <div class="export-options">
        <div class="panel" style="display:grid;gap:12px">
          <h2>הקובץ לייצוא</h2>
          <label class="field">חודש
            <select class="input" id="x-month">
              <option value="all">כל החודשים</option>
              ${months.map(m => `<option value="${m}" ${o.month === m ? 'selected' : ''}>${monthLabel(m, true)}</option>`).join('')}
            </select></label>
          <div class="check-list">
            ${sec('assigned', 'שובץ')}
            ${sec('open', 'לא שובץ')}
            ${sec('pool', 'מדריכים שעוד לא שובצו')}
            ${sec('courses', 'קורסים (לפי קורס)')}
            ${sec('done', 'הדרכות שבוצעו (תאריכים שעברו)')}
            <label><input type="checkbox" id="x-location" ${o.location ? 'checked' : ''}> כולל מיקום (מעבדה/בניין)</label>
          </div>
          <p class="small muted" style="margin:0">הסינון לפי תחום וסוג משבצת נלקח מלשונית השיבוץ.</p>
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            <button class="btn btn-primary" data-act="docx">הורדת קובץ Word</button>
            <button class="btn" data-act="copy-doc">העתקת הטקסט</button>
          </div>
        </div>
        <div class="panel" style="display:grid;gap:10px">
          <label class="field">פערים לטיפול (מופיע בראש הקובץ, שורה לכל פריט)
            <textarea class="input" id="x-gaps" rows="5" placeholder="למשל: למצוא מחליף להדר ב-29.10">${esc(state.notes.gaps)}</textarea></label>
          <label class="field">משימות (מופיע בסוף הקובץ)
            <textarea class="input" id="x-tasks" rows="4">${esc(state.notes.tasks)}</textarea></label>
        </div>
      </div>
      <article class="paper" id="x-preview" dir="rtl">${toHtml(exportBlocks())}</article>
    </div>`;
}

function refreshPreview() {
  const p = $('#x-preview');
  if (p) p.innerHTML = toHtml(exportBlocks());
}

// ---------------------------------------------------------------- files view

function renderFiles() {
  const el = $('#view-files');
  const r = state.imports.roster;
  const sch = state.imports.schedule || [];
  const lastSch = sch[sch.length - 1];
  const fmt = iso => new Date(iso).toLocaleString('he-IL', { dateStyle: 'short', timeStyle: 'short' });
  el.innerHTML = `
    <div class="files-grid">
      <div class="panel" style="display:grid;gap:10px">
        <h2>רשימת המדריכים</h2>
        <p class="small muted" style="margin:0">הקובץ "נוער שוחר מדע - שיבוץ מדריכים" (Excel). נקראים: שם, טלפון, הדרכה, הדרכת בטיחות, הערות (קוסם/ת, הרצאות, אילוצי זמן).</p>
        <label class="drop" data-drop="roster">
          <strong>גררו לכאן קובץ Excel</strong><span class="small muted">או לחצו לבחירה (‎.xlsx, ‎.xls, ‎.csv)</span>
          <input type="file" id="file-roster" accept=".xlsx,.xls,.csv" hidden>
        </label>
        <div class="status-line">${r ? `יובא: ${esc(r.name)} · ${r.count} מדריכים · ${fmt(r.at)}` : 'עוד לא יובא קובץ'}</div>
        ${notices.roster ? `<div class="result ${notices.roster.error ? 'error' : ''}">${esc(notices.roster.text)}</div>` : ''}
      </div>
      <div class="panel" style="display:grid;gap:10px">
        <h2>קובץ הלו"ז</h2>
        <p class="small muted" style="margin:0">"קובץ לוז" (שם השנה לא משנה) כ-PDF, כפי שנשמר מ-Excel: שבועות של 7 עמודות עם שורת תאריכים. אפשר לייבא שוב כשהלו"ז מתעדכן - השיבוצים שלכם נשמרים.</p>
        <label class="drop" data-drop="schedule">
          <strong>גררו לכאן קובץ PDF</strong><span class="small muted">או לחצו לבחירה</span>
          <input type="file" id="file-schedule" accept=".pdf,application/pdf" hidden>
        </label>
        <div class="status-line">${lastSch ? `יובא: ${esc(lastSch.name)} · ${lastSch.events} אירועים (${shortDate(lastSch.from)}–${shortDate(lastSch.to)}) · ${fmt(lastSch.at)}` : 'עוד לא יובא קובץ'}</div>
        ${notices.schedule ? `<div class="result ${notices.schedule.error ? 'error' : ''}">${esc(notices.schedule.text)}</div>` : ''}
      </div>
    </div>
    <div class="panel" style="display:grid;gap:10px">
      <h2>שמירה וגיבוי</h2>
      <p class="small muted" style="margin:0">${store?.kind === 'db' ? 'הנתונים נשמרים בחשבון שלכם ב-claude.ai, ורק אתם רואים אותם.' : 'הנתונים נשמרים בדפדפן הזה בלבד. כדאי להוריד גיבוי מדי פעם, ולהעביר איתו את העבודה למחשב אחר.'}</p>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn" data-act="backup" ${demo ? 'disabled' : ''}>הורדת גיבוי</button>
        <label class="btn">שחזור מגיבוי<input type="file" id="file-backup" accept=".json,application/json" hidden></label>
        <button class="btn btn-danger" data-act="reset" ${demo ? 'disabled' : ''}>מחיקת כל הנתונים</button>
      </div>
    </div>`;
}

// ---------------------------------------------------------------- imports

let pdfjsPromise = null;
function loadPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import(new URL('../vendor/pdf.min.js', import.meta.url).href).then(lib => {
      lib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.js', import.meta.url).href;
      return lib;
    }).catch(e => { pdfjsPromise = null; throw e; });
  }
  return pdfjsPromise;
}

async function importRoster(file) {
  notices.roster = { text: 'קורא את הקובץ…' };
  render();
  try {
    const XLSX = await loadScript(XLSX_URL, 'XLSX');
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    let rows = null;
    for (const name of wb.SheetNames) {
      const r = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
      if (r.some(row => row.some(c => /שם/.test(String(c))))) { rows = r; break; }
    }
    if (!rows) throw new Error('לא נמצא גיליון עם עמודת "שם"');
    const list = parseRosterRows(rows);
    if (!list.length) throw new Error('לא נמצאו מדריכים בקובץ');
    leaveDemo();
    state.instructors = list;
    state.imports.roster = { name: file.name, at: new Date().toISOString(), count: list.length };
    personIndex = null;
    // Names written in the schedule take the roster's spelling.
    const matcher = createNameMatcher(people().map(p => p.name));
    for (const ev of state.events) for (const s of ev.slots) for (const g of s.groups) for (const p of g.positions) {
      if (p.name) p.name = matcher.match(p.name) || p.name;
    }
    const roles = k => list.filter(i => i.roles[k] && i.active).length;
    const constraints = list.filter(i => i.availability.recognized).length;
    notices.roster = { text: `נקראו ${list.length} מדריכים: ${roles('lab')} מדריכים, ${roles('magic')} קוסמים, ${roles('lecture')} מרצים. אילוצי זמינות זוהו אצל ${constraints}. כדאי לעבור על לשונית "מדריכים" ולתקן מה שצריך.` };
    commit();
  } catch (e) {
    notices.roster = { error: true, text: 'הייבוא נכשל: ' + (e.message || e) };
    render();
  }
}

async function importSchedule(file) {
  notices.schedule = { text: 'קורא את הלו"ז… (כמה שניות)' };
  render();
  try {
    const pdfjs = await loadPdfjs();
    const { cells, warnings } = await extractScheduleCells(pdfjs, new Uint8Array(await file.arrayBuffer()));
    if (!cells.length) throw new Error('לא נמצאו תאים עם תאריכים. האם זה קובץ הלו"ז השבועי?');
    leaveDemo();
    const parsed = parseCells(cells, createNameMatcher(people().map(p => p.name)));
    const result = mergeSchedule(state.events, parsed);
    state.events = result.events;
    const dates = parsed.map(e => e.date).sort();
    state.imports.schedule = [...(state.imports.schedule || []).slice(-9), { name: file.name, at: new Date().toISOString(), events: parsed.length, from: dates[0], to: dates[dates.length - 1] }];
    const relevant = parsed.filter(e => isRelevant(e)).length;
    let text = `נמצאו ${parsed.length} תאים, מהם ${relevant} ימי עיון וקורסים עם משבצות לשיבוץ. חדשים: ${result.added}, עודכנו: ${result.updated}${result.removed ? `, הוסרו: ${result.removed}` : ''}.`;
    if (result.kept.length) text += ` ${result.kept.length} אירועים ששובצו בעבר לא מופיעים בקובץ החדש ומסומנים בלוח.`;
    if (warnings.length) text += ' ' + warnings.join(' ');
    notices.schedule = { text };
    // Start from the first month of the new file.
    const firstMonth = dates[0]?.slice(0, 7);
    if (firstMonth && state.settings.month !== 'all' && state.settings.month !== firstMonth) state.settings.month = firstMonth;
    commit();
  } catch (e) {
    console.error(e);
    notices.schedule = { error: true, text: 'הייבוא נכשל: ' + (e.message || e) };
    render();
  }
}

async function restoreBackup(file) {
  try {
    const data = JSON.parse(await file.text());
    if (!Array.isArray(data.events) || !Array.isArray(data.instructors)) throw new Error('זה לא קובץ גיבוי של האתר');
    leaveDemo();
    state = withDefaults(data);
    toast('הגיבוי שוחזר');
    commit();
  } catch (e) {
    toast('השחזור נכשל: ' + (e.message || e));
  }
}

// ---------------------------------------------------------------- render & routing

function render() {
  for (const v of ['board', 'people', 'export', 'files']) {
    $(`#view-${v}`).hidden = v !== view;
    $(`#tab-${v}`).toggleAttribute('aria-current', v === view);
    if (v === view) $(`#tab-${v}`).setAttribute('aria-current', 'page');
  }
  if (view === 'board') renderBoard();
  else if (view === 'people') renderPeople();
  else if (view === 'export') renderExport();
  else renderFiles();
}

function go(v) {
  view = v;
  render();
  window.scrollTo({ top: 0 });
}

// ---------------------------------------------------------------- events

document.addEventListener('click', async e => {
  const t = e.target;
  const goBtn = t.closest('[data-go]');
  if (goBtn) { go(goBtn.dataset.go); return; }
  const tab = t.closest('.tab');
  if (tab) { go(tab.dataset.view); return; }
  if (t.closest('[data-close]')) { closeDialog(); return; }

  // Dialog actions
  if (t.closest('#dialog')) {
    const pick = t.closest('[data-pick]');
    if (pick && dialogHandlers.pick) { dialogHandlers.pick(pick.dataset.pick); return; }
    const gact = t.closest('[data-gact]');
    if (gact && dialogHandlers.group) {
      const row = gact.closest('[data-gi]');
      dialogHandlers.group(gact.dataset.gact, row ? +row.dataset.gi : -1);
      return;
    }
    const act = t.closest('[data-act]')?.dataset.act;
    if (act === 'pick-free') { const v = $('#pick-free').value.trim(); if (v) dialogHandlers.pick(v); return; }
    if (act === 'pick-clear') { dialogHandlers.clear(); return; }
    if (act === 'slot-save' || act === 'event-save' || act === 'note-save' || act === 'avail-save' || act === 'person-save') { dialogHandlers.save(); return; }
    if (act === 'slot-delete') { dialogHandlers.remove(); return; }
    if (act === 'avail-reset') { dialogHandlers.reset(); return; }
    if (t === dialog()) closeDialog(); // backdrop click
    return;
  }

  const pos = t.closest('[data-pos]');
  if (pos) {
    const ev = findEvent(pos.closest('[data-ev]').dataset.ev);
    const [slotId, gi, pi] = pos.dataset.pos.split('|');
    const slot = findSlot(ev, slotId);
    if (slot) openPicker(ev, slot, +gi, +pi);
    return;
  }

  const actEl = t.closest('[data-act]');
  if (!actEl) return;
  const act = actEl.dataset.act;
  const evEl = actEl.closest('[data-ev]');
  const ev = evEl ? findEvent(evEl.dataset.ev) : null;

  switch (act) {
    case 'expand': actEl.classList.toggle('clamp'); break;
    case 'raw': evEl.querySelector('.raw').hidden = !evEl.querySelector('.raw').hidden; break;
    case 'note': noteEditor(ev); break;
    case 'hide': ev.ignored = true; commit(); toast('האירוע הוסתר. אפשר להחזיר אותו דרך "הצג מוסתרים".'); break;
    case 'unhide':
      if (ev.ignored) ev.ignored = false;
      else { ev.kind = ev.kind === 'other' ? 'study' : ev.kind; ev.kindOverride = ev.kind; }
      commit();
      if (!ev.slots.length) slotEditor(ev, newSlot(ev), true);
      break;
    case 'add-slot': slotEditor(ev, newSlot(ev), true); break;
    case 'edit-slot': {
      const slot = findSlot(ev, actEl.closest('[data-slot]').dataset.slot);
      if (slot) slotEditor(ev, slot, false);
      break;
    }
    case 'add-event': eventEditor(); break;
    case 'auto': {
      const snapshot = clone(state.events);
      const targets = state.events.filter(x => eventShown(x) && !x.ignored && isRelevant(x));
      const n = autoFill({ instructors: people(), events: state.events, targets, settings: state.settings });
      if (n) { autoUndo = snapshot; toast(`שובצו ${n} עמדות. העמדות בכחול שובצו אוטומטית - עברו עליהן.`); }
      else toast('לא נמצאו מדריכים פנויים לעמדות החסרות');
      commit();
      break;
    }
    case 'undo-auto':
      if (autoUndo) { state.events = autoUndo; autoUndo = null; toast('השיבוץ האוטומטי בוטל'); commit(); }
      break;
    case 'add-person': personEditor(); break;
    case 'edit-avail': {
      const p = people().find(x => x.id === actEl.closest('[data-person]').dataset.person);
      if (p) availabilityEditor(p);
      break;
    }
    case 'remind': {
      const p = people().find(x => x.id === actEl.closest('[data-person]').dataset.person);
      if (p) copyText(reminderText(p));
      break;
    }
    case 'docx': {
      actEl.disabled = true;
      try {
        const docx = await loadScript(DOCX_URL, 'docx');
        const title = exportTitle();
        const blob = await toDocxBlob(docx, exportBlocks(), title);
        const res = await saveFile(`${title}.docx`, blob);
        if (res === 'saved') toast('הקובץ נשמר');
      } catch (err) {
        toast('יצירת הקובץ נכשלה: ' + (err.message || err.code || err));
      } finally { actEl.disabled = false; }
      break;
    }
    case 'copy-doc': {
      const blocks = exportBlocks();
      copyText(toPlainText(blocks), toHtml(blocks));
      break;
    }
    case 'backup': {
      const date = todayISO();
      try { await saveFile(`גיבוי שיבוץ מדריכים ${date}.json`, JSON.stringify(state)); } catch (err) { toast('ההורדה נכשלה: ' + (err.message || err.code)); }
      break;
    }
    case 'reset':
      if (!actEl.classList.contains('armed')) {
        actEl.classList.add('armed');
        actEl.textContent = 'לחצו שוב למחיקה סופית';
        setTimeout(() => { if (actEl.isConnected) { actEl.classList.remove('armed'); actEl.textContent = 'מחיקת כל הנתונים'; } }, 4000);
      } else {
        state = emptyState();
        autoUndo = null;
        notices.roster = notices.schedule = null;
        toast('כל הנתונים נמחקו');
        commit();
      }
      break;
    default: break;
  }
});

document.addEventListener('change', e => {
  const t = e.target;
  const s = state.settings;
  if (t.dataset.filter) { s[t.dataset.filter][t.value] = t.checked; commit(); return; }
  if (t.dataset.toggle) { s[t.dataset.toggle] = t.checked; commit(); return; }
  if (t.dataset.setting === 'month') { s.month = t.value; commit(); return; }

  const personRow = t.closest('[data-person]');
  if (personRow) {
    const p = people().find(x => x.id === personRow.dataset.person);
    if (!p) return;
    if (t.dataset.role) editPerson(p.id, { roles: { ...p.roles, [t.dataset.role]: t.checked } });
    else if (t.hasAttribute('data-active')) editPerson(p.id, { active: t.checked });
    else if (t.hasAttribute('data-safety')) editPerson(p.id, { safetyOk: t.checked });
    commit();
    return;
  }

  if (t.id === 'x-month') { state.exportOptions.month = t.value; commit(); return; }
  if (t.id === 'x-location') { state.exportOptions.location = t.checked; commit(); return; }
  if (t.dataset.section) { state.exportOptions.sections[t.dataset.section] = t.checked; commit(); return; }

  if (t.id === 'file-roster' && t.files[0]) importRoster(t.files[0]);
  if (t.id === 'file-schedule' && t.files[0]) importSchedule(t.files[0]);
  if (t.id === 'file-backup' && t.files[0]) restoreBackup(t.files[0]);
});

let notesTimer = null;
document.addEventListener('input', e => {
  const t = e.target;
  if (t.id === 'pick-search') {
    const q = normalizeName(t.value);
    document.querySelectorAll('#pick-body .cand').forEach(c => { c.hidden = q && !c.dataset.search.includes(q); });
    if (q) document.querySelectorAll('#pick-body details').forEach(d => { d.open = true; });
    return;
  }
  if (t.id === 'people-search') {
    peopleQuery = t.value;
    const pos = t.selectionStart;
    renderPeople();
    const again = $('#people-search');
    again.focus();
    again.setSelectionRange(pos, pos);
    return;
  }
  if (t.id === 'x-gaps' || t.id === 'x-tasks') {
    state.notes[t.id === 'x-gaps' ? 'gaps' : 'tasks'] = t.value;
    clearTimeout(notesTimer);
    notesTimer = setTimeout(() => {
      refreshPreview();
      if (!demo) { dirty = true; flush(); }
    }, 400);
  }
});

document.addEventListener('keydown', e => {
  if (e.key !== 'Enter') return;
  if (e.target.id === 'pick-search') {
    const first = [...document.querySelectorAll('#pick-body .cand')].find(c => !c.hidden);
    if (first) { e.preventDefault(); dialogHandlers.pick(first.dataset.pick); }
  } else if (e.target.id === 'pick-free') {
    const v = e.target.value.trim();
    if (v) { e.preventDefault(); dialogHandlers.pick(v); }
  }
});

// Drag and drop onto the file zones.
document.addEventListener('dragover', e => {
  const zone = e.target.closest?.('[data-drop]');
  if (!zone) return;
  e.preventDefault();
  zone.classList.add('over');
});
document.addEventListener('dragleave', e => {
  const zone = e.target.closest?.('[data-drop]');
  if (zone) zone.classList.remove('over');
});
document.addEventListener('drop', e => {
  const zone = e.target.closest?.('[data-drop]');
  if (!zone) return;
  e.preventDefault();
  zone.classList.remove('over');
  const file = e.dataTransfer.files[0];
  if (!file) return;
  if (zone.dataset.drop === 'roster') importRoster(file);
  else importSchedule(file);
});

$('#dialog').addEventListener('close', () => { dialogHandlers = {}; });
document.addEventListener('toggle', e => { if (e.target.classList?.contains('filters-box')) filtersOpen = e.target.open; }, true);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });

// ---------------------------------------------------------------- start

async function start() {
  render();
  store = await openStore();
  let saved = null;
  try { saved = await store.load(); } catch { saved = null; }
  if (saved && hasData(withDefaults(saved))) {
    state = withDefaults(saved);
  } else {
    const d = demoState();
    state = { ...emptyState(), instructors: d.instructors, events: d.events };
    demo = true;
    $('#demo-banner').hidden = false;
  }
  personIndex = null;
  render();
}

start();
