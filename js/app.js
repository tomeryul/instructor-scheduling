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
import { icon } from './icons.js';

const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const DOCX_URL = 'https://cdn.jsdelivr.net/npm/docx@9.0.2/build/index.umd.js';

const DEFAULT_SETTINGS = {
  subjects: { chem: true, general: true, bio: false, cs: false, other: false },
  kinds: { study: true, course: true },
  slotKinds: { lab: true, lecture: true, magic: true },
  month: 'all',
  onlyOpen: false,
  onlyChanged: false,
  showHidden: false,
  showPast: false,
};
const DEFAULT_EXPORT = { month: 'all', location: false, sections: { assigned: true, open: true, pool: true, courses: true, done: false } };
const ROLE_LABELS = { lab: 'מדריך/ה', magic: 'קוסם/ת', lecture: 'מרצה' };
const VIEW_TITLES = { board: 'שיבוץ', people: 'מדריכים', export: 'ייצוא', files: 'קבצים' };
const SLOT_ICONS = { lab: 'flask.fill', lecture: 'mic.fill', magic: 'sparkles' };
const SLOT_ROLE = { lab: 'מדריך/ה', lecture: 'מרצה', magic: 'קוסם/ת' };

const $ = sel => document.querySelector(sel);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const clone = o => JSON.parse(JSON.stringify(o));
const chevron = () => icon('chevron.forward', 'chevron');

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
    lastImport: null,
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

const initials = name => name.split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('');

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
}

// ---------------------------------------------------------------- feedback

let toastTimer = null;
function toast(msg) {
  const el = $('#toast');
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, 3400);
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
    openSheet({
      title: 'העתקה',
      sub: 'ההעתקה האוטומטית לא זמינה כאן. הטקסט מסומן - העתיקו אותו.',
      body: `<div class="list"><div class="row"><textarea class="field-inline" id="copy-area" rows="14" readonly>${esc(text)}</textarea></div></div>`,
    });
    const ta = $('#copy-area');
    ta.focus();
    ta.select();
  }
}

// ---------------------------------------------------------------- sheets, alerts, menus

const dialog = () => $('#dialog');
let dialogHandlers = {};

/**
 * A sheet: grabber, Cancel (✕) on the leading edge, the task's title, Done (✓) on
 * the trailing edge when the sheet collects input, then optional search and body.
 */
function openSheet({ title, sub = '', done = false, search = '', body = '' }, handlers = {}) {
  const d = dialog();
  d.innerHTML = `
    <div class="sheet-inner">
      <div class="sheet-grabber" aria-hidden="true"></div>
      <div class="sheet-bar">
        <button class="icon-btn circle" data-close aria-label="ביטול" title="ביטול">${icon('xmark')}</button>
        <h2 id="dialog-title">${esc(title)}</h2>
        ${done ? `<button class="icon-btn circle prominent" data-act="sheet-done" aria-label="${esc(done)}" title="${esc(done)}">${icon('checkmark')}</button>` : '<span></span>'}
      </div>
      ${sub ? `<p class="sheet-sub">${sub}</p>` : ''}
      ${search ? `<div class="sheet-search">${search}</div>` : ''}
      <div class="sheet-body">${body}</div>
    </div>`;
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

let alertActions = [];
/** An alert for rare, irreversible actions: title, message, Cancel and the action. */
function openAlert({ title, message, actions }) {
  const a = $('#alert');
  alertActions = actions;
  a.innerHTML = `
    <h2 id="alert-title">${esc(title)}</h2>
    ${message ? `<p>${esc(message)}</p>` : ''}
    <div class="alert-actions">${actions.map((x, i) => `<button class="btn ${x.cls || 'btn-gray'}" data-alert="${i}">${esc(x.label)}</button>`).join('')}</div>`;
  a.showModal();
  a.focus();
}

let menuHandlers = null;
let menuScrollY = 0;
/** A pull-down menu next to the button that opened it. */
function openMenu(anchor, items) {
  const m = $('#menu');
  menuHandlers = items;
  m.innerHTML = items.map((it, i) => (it === '-' ? '<hr>' : `<button role="menuitem" data-menu="${i}" class="${it.destructive ? 'destructive' : ''}">${esc(it.label)}${icon(it.icon)}</button>`)).join('');
  m.hidden = false;
  const r = anchor.getBoundingClientRect();
  const w = m.offsetWidth, h = m.offsetHeight;
  let top = r.bottom + 6;
  if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 6);
  const left = Math.min(Math.max(8, r.left), window.innerWidth - w - 8);
  m.style.top = `${top}px`;
  m.style.left = `${left}px`;
  menuScrollY = window.scrollY;
  m.querySelector('button')?.focus({ preventScroll: true });
}

function closeMenu() {
  $('#menu').hidden = true;
  menuHandlers = null;
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
  if (s.onlyChanged && !(ev.changes || []).length) return false;
  return true;
}

/** Changes from the last schedule update that matter here: relevant, not hidden, in a subject in view. */
function relevantChanges() {
  const li = state.lastImport;
  if (!li || !li.report) return [];
  const byId = new Map(state.events.map(e => [e.id, e]));
  return li.report.filter(r => {
    const ev = byId.get(r.evId);
    return ev && !ev.ignored && isRelevant(ev) && visibleSlots(ev, state.settings).length;
  });
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

// ---------------------------------------------------------------- shared pieces

const titleBlock = (title, subtitle) => `
  <div class="title-block"><h1 class="large-title">${esc(title)}</h1>${subtitle ? `<p class="subtitle">${subtitle}</p>` : ''}</div>`;

function demoNotice() {
  if (!demo) return '';
  return `
    <div class="list"><div class="notice">
      <span class="row-icon">${icon('info.circle')}</span>
      <div class="row-main"><span class="row-title"><b>נתוני דוגמה</b></span>
        <span class="row-sub">השמות והאירועים כאן מומצאים, כדי להראות איך זה עובד. שינויים בדוגמה לא נשמרים.</span></div>
      <button class="btn btn-tinted btn-small" data-go="files" id="demo-import">ייבוא הקבצים שלי</button>
    </div></div>`;
}

const toggleChip = (attrs, label, checked, extra = '') => `
  <label class="toggle"><input type="checkbox" ${attrs} ${checked ? 'checked' : ''}>${icon('checkmark', 'check')}${extra}${esc(label)}</label>`;

const popup = (id, options, value, attrs = '') => `
  <span class="popup"><select id="${id}" ${attrs}>${options.map(([v, l]) => `<option value="${esc(v)}" ${v === value ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>${icon('chevron.updown')}</span>`;

// ---------------------------------------------------------------- board view

function renderBoard() {
  const el = $('#view-board');
  if (!hasData(state)) {
    el.innerHTML = `${titleBlock('שיבוץ')}
      <div class="list empty-state">${icon('calendar.fill')}
        <h2>מתחילים מהקבצים</h2>
        <p>ייבאו את רשימת המדריכים (Excel) ואת קובץ הלו"ז (PDF). האתר מוצא את ימי העיון, הקורסים, ההרצאות ומופעי הקסמים, ומשבצים מדריכים בלחיצה.</p>
        <button class="btn btn-prominent" data-go="files">לייבוא הקבצים</button>
      </div>`;
    return;
  }
  const s = state.settings;
  const today = todayISO();
  const months = monthsInData();
  const assignments = allAssignments(state.events);
  const shown = state.events.filter(eventShown);

  // Counts for the filter toggles, over the month in view.
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

  const count = n => ` <span class="count">${n}</span>`;
  const filterToggle = (group, key, label, n, color) =>
    toggleChip(`data-filter="${group}" value="${key}"`, label, s[group][key] !== false, (color ? `<span class="dot" style="background:${color}"></span>` : '')).replace('</label>', `${n !== undefined ? count(n) : ''}</label>`);
  const optionToggle = (key, label, n) => toggleChip(`data-toggle="${key}"`, label, s[key]).replace('</label>', `${n !== undefined ? count(n) : ''}</label>`);

  const days = [];
  for (const ev of shown) {
    const last = days[days.length - 1];
    if (last && last.date === ev.date) last.events.push(ev);
    else days.push({ date: ev.date, events: [ev] });
  }
  const subjectsOn = Object.entries(SUBJECTS).filter(([k]) => s.subjects[k] !== false).map(([, v]) => v).join(', ');
  const monthText = s.month === 'all' ? 'כל החודשים' : monthLabel(s.month, true);

  el.innerHTML = `
    ${titleBlock('שיבוץ', `${esc(monthText)} · ${open ? `${open} עמדות פתוחות` : 'הכול משובץ'}`)}
    ${demoNotice()}
    <div class="list summary-card">
      <div class="stats">
        <div class="stat open"><b>${open}</b><span>עמדות פתוחות</span></div>
        <div class="stat done"><b>${total - open}</b><span>משובצות</span></div>
        <div class="stat"><b>${shown.length}</b><span>אירועים</span></div>
      </div>
      <div class="progress" role="progressbar" aria-label="אחוז העמדות המשובצות" aria-valuenow="${pct}" aria-valuemin="0" aria-valuemax="100"><span style="width:${pct}%"></span></div>
      <div class="toolbar">
        <button class="btn btn-prominent" data-act="auto" ${open ? '' : 'disabled'}>${icon('wand')}שיבוץ אוטומטי של החסרים</button>
        ${autoUndo ? `<button class="btn btn-gray" data-act="undo-auto">${icon('arrow.uturn')}ביטול השיבוץ האוטומטי</button>` : ''}
      </div>
    </div>
    ${renderChanges()}
    <div class="list-section">
      <div class="list">
        <details class="disclosure" ${filtersOpen ? 'open' : ''} id="filters-box">
          <summary class="row"><span class="row-icon">${icon('slider')}</span>
            <span class="row-main"><span class="row-title">סינון</span><span class="row-sub">${esc(monthText)} · ${esc(subjectsOn || 'אין תחומים')}${s.onlyOpen ? ' · רק עם חסרים' : ''}</span></span>${chevron()}</summary>
          <div class="filters">
            <div class="filter-line"><span class="label">חודש</span>${popup('f-month', [['all', 'כל החודשים'], ...months.map(m => [m, monthLabel(m, true)])], s.month, 'data-setting="month"')}</div>
            <div class="filter-line"><span class="label">סוג</span>
              ${filterToggle('kinds', 'study', 'ימי עיון', kindCounts.study || 0, 'var(--yellow)')}
              ${filterToggle('kinds', 'course', 'קורסים', kindCounts.course || 0, 'var(--purple)')}</div>
            <div class="filter-line"><span class="label">משבצות</span>
              ${filterToggle('slotKinds', 'lab', 'הדרכות', slotKindCounts.lab || 0, 'var(--icon-lab)')}
              ${filterToggle('slotKinds', 'lecture', 'הרצאות', slotKindCounts.lecture || 0, 'var(--icon-lecture)')}
              ${filterToggle('slotKinds', 'magic', 'מופעי קסמים', slotKindCounts.magic || 0, 'var(--icon-magic)')}</div>
            <div class="filter-line"><span class="label">תחום</span>
              ${Object.entries(SUBJECTS).map(([k, label]) => filterToggle('subjects', k, label, subjectCounts[k] || 0)).join('')}</div>
            <div class="filter-line"><span class="label">הצגה</span>
              ${optionToggle('onlyOpen', 'רק עם חסרים')}
              ${optionToggle('showPast', 'כולל שעברו', pastCount)}
              ${optionToggle('showHidden', 'מוסתרים ולא רלוונטיים', hiddenCount)}</div>
          </div>
        </details>
      </div>
      <div class="legend list-footer">
        <span><span class="pos empty">${icon('plus')}חסר</span></span>
        <span><span class="pos">משובץ</span></span>
        <span><span class="pos src-auto">${icon('sparkles')}אוטומטי</span></span>
        <span><span class="pos warn">${icon('exclamation.triangle.fill')}הסתייגות</span></span>
        <span><span class="pos bad">${icon('xmark.circle.fill')}לא זמין/ה או כפילות</span></span>
      </div>
    </div>
    ${days.length ? days.map(d => `
      <section class="list-section day-group">
        <h2 class="list-header day">יום ${DAY_NAMES[weekday(d.date)]} <span class="secondary">${shortDate(d.date)}${d.date < today ? ' · עבר' : ''}</span></h2>
        ${d.events.map(ev => renderEvent(ev, assignments)).join('')}
      </section>`).join('') : `
      <div class="list empty-state">${icon('magnifyingglass')}<h2>אין אירועים להצגה</h2><p>אין אירועים שמתאימים לסינון. אפשר לשנות את החודש או את התחומים בסינון.</p></div>`}`;
}

function renderChanges() {
  const li = state.lastImport;
  if (!li || li.dismissed) return '';
  const items = relevantChanges();
  const when = new Date(li.at).toLocaleString('he-IL', { weekday: 'long', hour: '2-digit', minute: '2-digit' });
  const row = r => `
    <button class="row chg-${r.level}" data-jump="${esc(r.evId)}">
      <span class="row-icon">${icon(r.level === 'new' ? 'plus' : r.level === 'info' ? 'info.circle' : 'exclamation.triangle.fill')}</span>
      <span class="row-main"><span class="row-title">${esc(r.level === 'new' ? r.title.split('(')[0].trim().slice(0, 70) : r.text)}</span>
        <span class="row-sub">${esc(dayLabel(r.date))}${r.level === 'new' ? '' : ` · ${esc(r.title.split('(')[0].trim().slice(0, 50))}`}</span></span>${chevron()}
    </button>`;
  const attention = items.filter(r => r.level === 'attention');
  const fresh = items.filter(r => r.level === 'new');
  const info = items.filter(r => r.level === 'info');
  return `
    <section class="list-section changes">
      <h2 class="list-header">מה השתנה בעדכון הלו"ז · ${esc(li.name)} · ${esc(when)}</h2>
      <div class="list changes-list">
        ${items.length ? '' : '<div class="row"><span class="row-main"><span class="row-sub">אין שינויים שנוגעים לשיבוצים או לאירועים שמוצגים כאן.</span></span></div>'}
        ${attention.map(row).join('')}
        ${fresh.map(row).join('')}
        ${info.length ? `<details class="disclosure"><summary class="row"><span class="row-icon">${icon('info.circle')}</span><span class="row-main"><span class="row-title">עדכונים שהלו"ז כבר מילא</span><span class="row-sub">${info.length} שינויים בשמות</span></span>${chevron()}</summary>${info.map(row).join('')}</details>` : ''}
      </div>
      <div class="toolbar list-footer">
        ${items.length ? toggleChip('data-toggle="onlyChanged"', 'רק אירועים שהשתנו', state.settings.onlyChanged) : ''}
        <button class="btn btn-plain btn-small" data-act="dismiss-changes">הסתרת הרשימה</button>
      </div>
    </section>`;
}

function renderEvent(ev, assignments) {
  const c = openCount(ev, state.settings);
  const kind = ev.kind || 'other';
  const slots = slotsShown(ev);
  const subjects = [...new Set(slots.map(sl => sl.subject).filter(Boolean))].map(k => SUBJECTS[k]);
  const hidden = ev.ignored || !isRelevant(ev);
  const changes = (ev.changes || []).filter(x => x.level !== 'new');
  const status = hidden ? '<span class="badge">מוסתר</span>'
    : c.open ? `<span class="badge badge-open">חסרים ${c.open}</span>`
      : c.total ? `<span class="badge badge-done">${icon('checkmark')}שובץ</span>` : '';
  return `
  <article class="list ev ${hidden ? 'is-ignored' : ''}" data-ev="${esc(ev.id)}">
    <div class="ev-head">
      <div class="row-main">
        <div class="badges">
          <span class="badge ${kind === 'study' ? 'badge-study' : kind === 'course' ? 'badge-course' : ''}">${EVENT_KINDS[kind]}</span>
          ${subjects.map(t => `<span class="badge">${esc(t)}</span>`).join('')}
          ${status}
          ${ev.missingFromSchedule ? `<span class="badge badge-warn" title="האירוע לא נמצא בקובץ הלו&quot;ז האחרון שיובא">${icon('exclamation.triangle.fill')}לא בלו"ז העדכני</span>`
            : changes.some(x => x.level === 'attention') ? `<span class="badge badge-warn">${icon('exclamation.triangle.fill')}השתנה</span>` : ''}
          ${(ev.changes || []).some(x => x.level === 'new') ? '<span class="badge badge-new">חדש</span>' : ''}
          ${ev.manual ? '<span class="badge">ידני</span>' : ''}
        </div>
        <h3 class="ev-title clamp" data-act="expand" title="לחצו להצגת הכותרת המלאה">${esc(ev.title || '(ללא כותרת)')}</h3>
      </div>
      <button class="icon-btn" data-act="ev-menu" aria-label="פעולות לאירוע" title="פעולות">${icon('ellipsis')}</button>
    </div>
    ${changes.map(x => `<div class="ev-change ${x.level}"><span class="row-icon">${icon(x.level === 'attention' ? 'exclamation.triangle.fill' : 'info.circle')}</span><span class="row-main">${esc(x.text)}</span></div>`).join('')}
    ${ev.note ? `<div class="ev-change note"><span class="row-icon">${icon('note')}</span><span class="row-main">${esc(ev.note)}</span></div>` : ''}
    ${slots.length ? slots.map(sl => renderSlot(ev, sl, assignments)).join('')
      : '<div class="row"><span class="row-main"><span class="row-sub">אין משבצות שדורשות מדריכים. אפשר להוסיף משבצת מתפריט הפעולות.</span></span></div>'}
    <pre class="ev-raw" hidden>${esc(ev.raw || '')}</pre>
  </article>`;
}

function renderSlot(ev, sl, assignments) {
  return `
  <div class="row slot slot-${sl.kind}" data-slot="${esc(sl.id)}">
    <span class="row-icon" title="${SLOT_KINDS[sl.kind]}">${icon(SLOT_ICONS[sl.kind])}</span>
    <div class="row-main">
      <div>
        <div class="row-title">${esc(sl.label)}</div>
        <div class="row-sub"><span class="slot-time">${esc(sl.start)}–${esc(sl.end)}</span> · ${SLOT_KINDS[sl.kind]}${sl.location ? ` · ${esc(sl.location)}` : ''}</div>
      </div>
      <div class="groups">
        ${sl.groups.map((g, gi) => `
          <div class="group">
            ${g.label ? `<span class="g-label">${esc(g.label)}${g.activity && g.activity !== sl.label ? ' · ' + esc(g.activity) : ''}</span>` : ''}
            ${g.positions.map((p, pi) => renderPosition(ev, sl, gi, pi, p, assignments)).join('')}
          </div>`).join('')}
      </div>
    </div>
    <button class="icon-btn" data-act="edit-slot" aria-label="פרטי המשבצת" title="פרטי המשבצת">${icon('info.circle')}</button>
  </div>`;
}

function renderPosition(ev, sl, gi, pi, p, assignments) {
  const ref = `data-pos="${esc(sl.id)}|${gi}|${pi}"`;
  if (!p.name) return `<button class="pos empty" ${ref} aria-label="בחירת ${SLOT_ROLE[sl.kind]}">${icon('plus')}${SLOT_ROLE[sl.kind]}</button>`;
  const issues = positionIssues(ev, sl, gi, pi, p.name, assignments);
  const cls = issues.level === 'bad' ? 'bad' : issues.level === 'warn' ? 'warn' : issues.level === 'external' ? 'external' : p.source === 'auto' ? 'src-auto' : '';
  const sym = issues.level === 'bad' ? icon('xmark.circle.fill') : issues.level === 'warn' ? icon('exclamation.triangle.fill') : p.source === 'auto' && issues.level === 'ok' ? icon('sparkles') : '';
  const title = [p.source === 'auto' ? 'שובץ אוטומטית' : p.source === 'pdf' ? 'מופיע בלו"ז' : 'שובץ ידנית', ...issues.reasons].join(' · ');
  return `<button class="pos ${cls}" ${ref} title="${esc(title)}">${sym}${esc(p.name)}</button>`;
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
    <button class="row cand tier-${c.tier}" data-pick="${esc(c.instructor.name)}" data-search="${esc(normalizeName(c.instructor.name))}">
      <span class="row-main"><span class="row-title">${esc(c.instructor.name)}</span>
        ${c.reasons.length || c.partner ? `<span class="row-sub">${c.partner ? '<span class="pair">עובדים יחד הרבה</span>' : ''}${c.reasons.map(r => `<span class="why">${esc(r)}</span>`).join('')}</span>` : ''}</span>
      <span class="row-value">${c.load} החודש</span>
      ${c.instructor.name === current ? `<span class="check-mark" aria-label="משובץ/ת עכשיו">${icon('checkmark')}</span>` : ''}
    </button>`;
  const sections = tiers.map(([tier, title], i) => {
    const list = ranked.filter(c => c.tier === tier);
    if (!list.length) return '';
    if (i < 2) return `<section class="list-section"><h3 class="list-header">${title} · ${list.length}</h3><div class="list">${list.map(row).join('')}</div></section>`;
    return `<section class="list-section"><div class="list"><details class="disclosure more">
      <summary class="row"><span class="row-main"><span class="row-title">${title}</span><span class="row-sub">${list.length} אנשים</span></span>${chevron()}</summary>
      ${list.map(row).join('')}</details></div></section>`;
  }).join('');
  const roleName = SLOT_ROLE[slot.kind];

  openSheet({
    title: `בחירת ${roleName}`,
    sub: `${esc(dayLabel(ev.date))} · <span class="slot-time">${esc(slot.start)}–${esc(slot.end)}</span> · ${esc(slot.label)}${group.label ? ' · ' + esc(group.label) : ''}`,
    search: '<input class="search" id="pick-search" type="search" placeholder="חיפוש שם" autocomplete="off" autofocus aria-label="חיפוש שם">',
    body: `<div id="pick-body" style="display:grid;gap:24px">${sections || '<div class="list empty-state"><p>אין מדריכים ברשימה. ייבאו את קובץ המדריכים, או כתבו שם למטה.</p></div>'}</div>
      <section class="list-section">
        <h3 class="list-header">שם שלא ברשימה</h3>
        <div class="list"><div class="row"><input class="field-inline" id="pick-free" placeholder="שם מלא" autocomplete="off" aria-label="שם שלא ברשימה"><button class="btn btn-tinted btn-small" data-act="pick-free">שיבוץ</button></div></div>
      </section>
      ${current ? `<div class="list"><button class="row action destructive" data-act="pick-clear">${icon('xmark.circle.fill')}<span class="row-main">ניקוי השיבוץ (${esc(current)})</span></button></div>` : ''}`,
  }, {
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
    <div class="row" data-gi="${gi}">
      <input class="field-inline" data-g="label" value="${esc(g.label)}" placeholder="קבוצה ${gi + 1} (שם לא חובה)" aria-label="שם הקבוצה">
      <span class="row-value">${g.positions.length} עמדות</span>
      <span class="stepper"><button data-gact="minus" aria-label="פחות עמדות" ${g.positions.length <= 1 ? 'disabled' : ''}>−</button><button data-gact="plus" aria-label="עוד עמדה">${icon('plus')}</button></span>
      ${draft.groups.length > 1 ? `<button class="icon-btn" data-gact="remove" aria-label="הסרת הקבוצה" style="color:var(--red)">${icon('trash')}</button>` : ''}
    </div>`).join('') + `<button class="row action" data-gact="add-group">${icon('plus.circle.fill')}<span class="row-main">הוספת קבוצה</span></button>`;
  const readForm = () => {
    draft.start = $('#se-start').value || draft.start;
    draft.end = $('#se-end').value || draft.end;
    draft.kind = $('#se-kind').value;
    draft.subject = $('#se-subject').value === 'general' ? null : $('#se-subject').value;
    draft.label = $('#se-label').value.trim() || SLOT_KINDS[draft.kind];
    draft.location = $('#se-location').value.trim();
    document.querySelectorAll('#se-groups [data-gi]').forEach(row => { draft.groups[+row.dataset.gi].label = row.querySelector('[data-g="label"]').value.trim(); });
  };
  openSheet({
    title: isNew ? 'משבצת חדשה' : 'פרטי המשבצת',
    sub: `${esc(dayLabel(ev.date))} · ${esc(ev.title.split('(')[0].trim().slice(0, 70))}`,
    done: 'שמירה',
    body: `
      <div class="list">
        <div class="row"><span class="row-label">התחלה</span><span class="row-main"></span><input type="time" id="se-start" value="${esc(draft.start)}"></div>
        <div class="row"><span class="row-label">סיום</span><span class="row-main"></span><input type="time" id="se-end" value="${esc(draft.end)}"></div>
        <div class="row"><span class="row-label">סוג</span><span class="row-main"></span>${popup('se-kind', Object.entries(SLOT_KINDS), draft.kind)}</div>
        <div class="row"><span class="row-label">תחום</span><span class="row-main"></span>${popup('se-subject', Object.entries(SUBJECTS), draft.subject || 'general')}</div>
      </div>
      <div class="list">
        <div class="row"><input class="field-inline" id="se-label" value="${esc(draft.label)}" placeholder="תיאור הפעילות" aria-label="תיאור הפעילות"></div>
        <div class="row"><input class="field-inline" id="se-location" value="${esc(draft.location || '')}" placeholder="מיקום" aria-label="מיקום"></div>
      </div>
      <section class="list-section">
        <h3 class="list-header">קבוצות ועמדות</h3>
        <div class="list" id="se-groups">${renderGroups()}</div>
        <p class="list-footer">בהדרכה יש בדרך כלל זוג מדריכים לכל קבוצה.</p>
      </section>
      ${isNew ? '' : `<div class="list"><button class="row action destructive" data-act="slot-delete">${icon('trash')}<span class="row-main">מחיקת המשבצת</span></button></div>`}`,
  }, {
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
  openSheet({
    title: 'אירוע ידני',
    sub: 'לאירוע שלא מופיע בקובץ הלו"ז. אחרי השמירה מוסיפים לו משבצת.',
    done: 'יצירה',
    body: `
      <div class="list">
        <div class="row"><span class="row-label">תאריך</span><span class="row-main"></span><input type="date" id="ee-date" value="${todayISO()}"></div>
        <div class="row"><span class="row-label">סוג</span><span class="row-main"></span>${popup('ee-kind', [['study', 'יום עיון'], ['course', 'קורס']], 'study')}</div>
      </div>
      <div class="list"><div class="row"><input class="field-inline" id="ee-title" placeholder="כותרת, למשל: יום עיון בכימיה לתיכון… כיתה י״א" aria-label="כותרת" autofocus></div></div>`,
  }, {
    save() {
      const date = $('#ee-date').value;
      const title = $('#ee-title').value.trim();
      if (!date || !title) { toast('צריך תאריך וכותרת'); return; }
      leaveDemo();
      const ev = { id: `m${Date.now().toString(36)}`, date, color: null, kind: $('#ee-kind').value, subject: null, title, raw: '', slots: [], ignored: false, note: '', manual: true, changes: [] };
      state.events.push(ev);
      state.events.sort((a, b) => a.date.localeCompare(b.date));
      closeDialog();
      commit();
      slotEditor(ev, newSlot(ev), true);
    },
  });
}

function noteEditor(ev) {
  openSheet({
    title: 'הערה',
    sub: `${esc(dayLabel(ev.date))} · ההערה מופיעה גם בקובץ המיוצא`,
    done: 'שמירה',
    body: `<div class="list"><div class="row"><textarea class="field-inline" id="note-text" rows="4" placeholder="למשל: תוזכרו ביום חמישי / השעות הוקדמו, כל המדריכים עודכנו" aria-label="הערה" autofocus>${esc(ev.note || '')}</textarea></div></div>`,
  }, {
    save() {
      ev.note = $('#note-text').value.trim();
      closeDialog();
      commit();
    },
  });
}

function eventMenu(anchor, ev) {
  const hidden = ev.ignored || !isRelevant(ev);
  const items = [
    { label: ev.note ? 'עריכת הערה' : 'הוספת הערה', icon: 'note', run: () => noteEditor(ev) },
    ...(ev.raw ? [{ label: 'הטקסט מהלו"ז', icon: 'text.doc', run: () => { const r = document.querySelector(`[data-ev="${CSS.escape(ev.id)}"] .ev-raw`); if (r) r.hidden = !r.hidden; } }] : []),
    { label: 'הוספת משבצת', icon: 'plus', run: () => slotEditor(ev, newSlot(ev), true) },
    '-',
    hidden
      ? { label: ev.ignored ? 'החזרה לשיבוץ' : 'סימון כרלוונטי', icon: 'eye', run: () => {
        if (ev.ignored) ev.ignored = false;
        else { ev.kind = ev.kind === 'other' ? 'study' : ev.kind; ev.kindOverride = ev.kind; }
        commit();
        if (!ev.slots.length) slotEditor(ev, newSlot(ev), true);
      } }
      : { label: 'הסתרה', icon: 'eye.slash', run: () => { ev.ignored = true; commit(); toast('האירוע הוסתר. אפשר להחזיר אותו דרך "מוסתרים" בסינון.'); } },
    ...(ev.missingFromSchedule || ev.manual ? [{ label: 'מחיקה', icon: 'trash', destructive: true, run: () => { state.events = state.events.filter(x => x !== ev); commit(); toast('האירוע נמחק'); } }] : []),
  ];
  openMenu(anchor, items);
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

function weekStrip(a) {
  return `<div class="week" aria-label="זמינות בשבוע">${[0, 1, 2, 3, 4, 5].map(d => {
    const rule = a?.days?.[d];
    const cls = rule === 'no' ? 'no' : Array.isArray(rule) && !(rule.length === 1 && rule[0][0] === '00:00' && rule[0][1] === '24:00') ? 'part' : '';
    return `<span class="${cls}" title="${DAY_NAMES[d]}: ${rule === 'no' ? 'לא פנוי/ה' : windowsToText(rule) || 'פנוי/ה'}">${'אבגדהו'[d]}'</span>`;
  }).join('')}</div>`;
}

function renderPeople() {
  const el = $('#view-people');
  const list = people();
  if (!list.length) {
    el.innerHTML = `${titleBlock('מדריכים')}
      <div class="list empty-state">${icon('person.2.fill')}<h2>אין עדיין מדריכים</h2>
        <p>ייבאו את קובץ המדריכים (Excel) בלשונית "קבצים", או הוסיפו מדריך/ה.</p>
        <button class="btn btn-prominent" data-go="files">לייבוא</button></div>`;
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
  const count = k => list.filter(p => p.active !== false && p.roles?.[k]).length;
  el.innerHTML = `
    ${titleBlock('מדריכים', `${list.length} ברשימה · ${count('lab')} מדריכים · ${count('magic')} קוסמים · ${count('lecture')} מרצים`)}
    ${demoNotice()}
    <input class="search" type="search" id="people-search" placeholder="חיפוש מדריך/ה" value="${esc(peopleQuery)}" aria-label="חיפוש מדריך/ה">
    <section class="list-section">
      <div class="list" id="people-list">
        ${rows.map(p => {
          const roles = Object.entries(ROLE_LABELS).filter(([k]) => p.roles?.[k]).map(([, v]) => v).join(' · ') || 'בלי תפקיד';
          const warn = [p.active === false ? 'לא פעיל/ה' : '', p.roles?.lab && !p.safetyOk ? 'בטיחות' : '', ...(p.tags || []).filter(t => /לא הגיע|מילואים/.test(t))].filter(Boolean);
          const avail = describeAvailability(p.availability);
          return `
          <button class="row" data-act="open-person" data-person="${esc(p.id)}">
            <span class="avatar ${p.active === false ? 'inactive' : ''}" aria-hidden="true">${esc(initials(p.name))}</span>
            <span class="row-main">
              <span class="row-title">${esc(p.name)}</span>
              <span class="row-sub">${esc(roles)}${avail ? ` · ${esc(avail)}` : ''}</span>
              ${warn.length ? `<span class="badges">${warn.map(w => `<span class="badge badge-warn">${icon('exclamation.triangle.fill')}${esc(w)}</span>`).join('')}</span>` : ''}
            </span>
            <span class="row-value" title="שיבוצים קרובים">${upcoming.get(p.id) || 0}</span>${chevron()}
          </button>`;
        }).join('') || '<div class="row"><span class="row-sub">לא נמצאו מדריכים בשם הזה.</span></div>'}
      </div>
      <p class="list-footer">התפקידים והזמינות נקראו מהעמודות "הדרכה" ו"הערות". המספר בצד הוא שיבוצים קרובים. תיקונים כאן נשמרים גם כשמייבאים את הקובץ מחדש.</p>
    </section>`;
}

/** Everything about one instructor: roles, status, weekly availability, reminder. */
function personSheet(person) {
  const a = person.availability || { days: {}, preferredDays: [] };
  const edited = !!state.instructorEdits[person.id]?.availability;
  openSheet({
    title: person.name,
    done: 'שמירה',
    body: `
      <div class="person-head">
        <span class="avatar big ${person.active === false ? 'inactive' : ''}" aria-hidden="true">${esc(initials(person.name))}</span>
        <h3>${esc(person.name)}</h3>
        ${person.phone ? `<span class="secondary num" style="direction:ltr;unicode-bidi:isolate;user-select:all">${esc(person.phone)}</span>` : ''}
        ${person.degree || person.tags?.length ? `<div class="person-tags">${person.degree ? `<span class="badge">${esc(person.degree)}</span>` : ''}${(person.tags || []).map(t => `<span class="badge">${esc(t)}</span>`).join('')}</div>` : ''}
        ${weekStrip(a)}
      </div>
      <section class="list-section">
        <h3 class="list-header">תפקידים</h3>
        <div class="list">
          ${Object.entries(ROLE_LABELS).map(([k, v]) => `<label class="row"><span class="row-main">${v}</span><input type="checkbox" role="switch" class="switch" data-prole="${k}" ${person.roles?.[k] ? 'checked' : ''}></label>`).join('')}
        </div>
      </section>
      <section class="list-section">
        <h3 class="list-header">מצב</h3>
        <div class="list">
          <label class="row"><span class="row-main">פעיל/ה</span><input type="checkbox" role="switch" class="switch" id="ps-active" ${person.active !== false ? 'checked' : ''}></label>
          <label class="row"><span class="row-main"><span class="row-title">הדרכת בטיחות הושלמה</span>${person.safetyNote ? `<span class="row-sub">${esc(person.safetyNote)}</span>` : ''}</span><input type="checkbox" role="switch" class="switch" id="ps-safety" ${person.safetyOk ? 'checked' : ''}></label>
        </div>
      </section>
      <section class="list-section">
        <h3 class="list-header">זמינות בשבוע</h3>
        <div class="list">
          ${[0, 1, 2, 3, 4, 5].map(d => `
            <div class="row"><span class="row-label">${DAY_NAMES[d]}</span>
              <input class="field-inline" data-day="${d}" value="${esc(windowsToText(a.days?.[d]))}" placeholder="פנוי/ה" aria-label="זמינות ביום ${DAY_NAMES[d]}">
              ${toggleChip(`data-pref="${d}"`, 'מועדף', a.preferredDays?.includes(d))}</div>`).join('')}
          <div class="row"><span class="row-label">לא זמין/ה עד</span><span class="row-main"></span><input type="date" id="av-until" value="${esc(a.unavailableUntil || '')}"></div>
          <div class="row"><span class="row-label">שעות מועדפות</span><input class="field-inline" id="av-hours" value="${a.preferredHours ? esc(a.preferredHours.join('-')) : ''}" placeholder="למשל 9:00-14:00"></div>
        </div>
        <p class="list-footer">ריק = פנוי/ה כל היום · "לא" = לא פנוי/ה · שעות: "מ-12:00", "עד 14:00", "10:00-13:45" (כמה, עם פסיק).</p>
      </section>
      ${person.notes ? `<section class="list-section"><h3 class="list-header">הערות מהקובץ</h3><div class="list"><div class="row"><span class="row-main" style="white-space:pre-wrap">${esc(person.notes)}</span></div></div></section>` : ''}
      <div class="list">
        <button class="row action" data-act="remind" data-person="${esc(person.id)}">${icon('bell')}<span class="row-main">העתקת הודעת תזכורת</span></button>
        ${edited ? `<button class="row action" data-act="avail-reset">${icon('arrow.uturn')}<span class="row-main">החזרת הזמינות למה שנקרא מהקובץ</span></button>` : ''}
      </div>`,
  }, {
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
      const availability = { days, preferredDays, preferredHours: hours[0] || null, unavailableUntil: $('#av-until').value || null, recognized: true };
      const roles = {};
      document.querySelectorAll('#dialog [data-prole]').forEach(i => { roles[i.dataset.prole] = i.checked; });
      const patch = { roles, active: $('#ps-active').checked, safetyOk: $('#ps-safety').checked };
      // Keep the availability read from the file unless it was changed here.
      if (JSON.stringify(availability) !== JSON.stringify({ ...person.availability, recognized: true })) patch.availability = availability;
      editPerson(person.id, patch);
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
  openSheet({
    title: 'מדריך/ה חדש/ה',
    done: 'הוספה',
    body: `
      <div class="list">
        <div class="row"><input class="field-inline" id="pe-name" placeholder="שם מלא" aria-label="שם" autofocus></div>
        <div class="row"><input class="field-inline" id="pe-phone" placeholder="טלפון" inputmode="tel" aria-label="טלפון"></div>
      </div>
      <section class="list-section">
        <h3 class="list-header">תפקידים</h3>
        <div class="list">${Object.entries(ROLE_LABELS).map(([k, v]) => `<label class="row"><span class="row-main">${v}</span><input type="checkbox" role="switch" class="switch" data-newrole="${k}" ${k === 'lab' ? 'checked' : ''}></label>`).join('')}</div>
      </section>`,
  }, {
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
  const sw = (attrs, label, checked) => `<label class="row"><span class="row-main">${label}</span><input type="checkbox" role="switch" class="switch" ${attrs} ${checked ? 'checked' : ''}></label>`;
  el.innerHTML = `
    ${titleBlock('ייצוא', 'קובץ Word במבנה של קובץ "משימות"')}
    ${demoNotice()}
    <div class="export-layout">
      <div class="export-side">
        <section class="list-section">
          <div class="list">
            <div class="row"><span class="row-label">חודש</span><span class="row-main"></span>${popup('x-month', [['all', 'כל החודשים'], ...months.map(m => [m, monthLabel(m, true)])], o.month)}</div>
          </div>
        </section>
        <section class="list-section">
          <h3 class="list-header">מה ייכנס לקובץ</h3>
          <div class="list">
            ${sw('data-section="assigned"', 'שובץ', o.sections.assigned !== false)}
            ${sw('data-section="open"', 'לא שובץ', o.sections.open !== false)}
            ${sw('data-section="pool"', 'מדריכים שעוד לא שובצו', o.sections.pool !== false)}
            ${sw('data-section="courses"', 'קורסים, לפי קורס', o.sections.courses !== false)}
            ${sw('data-section="done"', 'הדרכות שבוצעו', o.sections.done !== false)}
            ${sw('id="x-location"', 'כולל מיקום (מעבדה ובניין)', o.location)}
          </div>
          <p class="list-footer">התחומים וסוגי המשבצות נלקחים מהסינון בלשונית "שיבוץ".</p>
        </section>
        <div class="toolbar">
          <button class="btn btn-prominent" data-act="docx">${icon('square.and.arrow.down')}הורדת קובץ Word</button>
          <button class="btn btn-tinted" data-act="copy-doc">${icon('doc.on.doc')}העתקת הטקסט</button>
        </div>
        <section class="list-section">
          <h3 class="list-header">פערים לטיפול · בראש הקובץ, שורה לכל פריט</h3>
          <div class="list"><div class="row"><textarea class="field-inline" id="x-gaps" rows="4" placeholder="למשל: למצוא מחליף להדר ב-29.10" aria-label="פערים לטיפול">${esc(state.notes.gaps)}</textarea></div></div>
        </section>
        <section class="list-section">
          <h3 class="list-header">משימות · בסוף הקובץ</h3>
          <div class="list"><div class="row"><textarea class="field-inline" id="x-tasks" rows="3" aria-label="משימות">${esc(state.notes.tasks)}</textarea></div></div>
        </section>
      </div>
      <article class="paper" id="x-preview" dir="rtl" aria-label="תצוגה מקדימה של הקובץ">${toHtml(exportBlocks())}</article>
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
  const result = n => (n ? `<div class="result ${n.error ? 'error' : n.busy ? 'busy' : ''}">${esc(n.text)}</div>` : '');
  el.innerHTML = `
    ${titleBlock('קבצים', 'מכאן נכנסים הנתונים לאתר')}
    ${demoNotice()}
    <section class="list-section">
      <h3 class="list-header">רשימת המדריכים · Excel</h3>
      <div class="list">
        <label class="drop" data-drop="roster">${icon('tablecells')}
          <strong>גררו לכאן את קובץ המדריכים</strong><span class="secondary" style="font-size:.882rem">או לחצו לבחירה · ‎.xlsx ‎.xls ‎.csv</span>
          <input type="file" id="file-roster" accept=".xlsx,.xls,.csv" hidden>
        </label>
        ${result(notices.roster)}
      </div>
      <p class="list-footer">${r ? `יובא: ${esc(r.name)} · ${r.count} מדריכים · ${fmt(r.at)}` : 'נקראים: שם, טלפון, הדרכה, הדרכת בטיחות והערות (קוסם/ת, הרצאות, אילוצי זמן).'}</p>
    </section>
    <section class="list-section">
      <h3 class="list-header">קובץ הלו"ז · PDF</h3>
      <div class="list">
        <label class="drop" data-drop="schedule">${icon('pdf')}
          <strong>גררו לכאן את קובץ הלו"ז</strong><span class="secondary" style="font-size:.882rem">או לחצו לבחירה · כפי שנשמר מ-Excel</span>
          <input type="file" id="file-schedule" accept=".pdf,application/pdf" hidden>
        </label>
        ${result(notices.schedule)}
      </div>
      <p class="list-footer">${lastSch ? `יובא: ${esc(lastSch.name)} · ${lastSch.events} תאים (${shortDate(lastSch.from)}–${shortDate(lastSch.to)}) · ${fmt(lastSch.at)}. ` : ''}כשהלו"ז מתעדכן מייבאים אותו שוב - השיבוצים שלכם נשמרים.</p>
    </section>
    <section class="list-section">
      <h3 class="list-header">שמירה וגיבוי</h3>
      <div class="list">
        <button class="row action" data-act="backup" ${demo ? 'disabled' : ''}>${icon('square.and.arrow.down')}<span class="row-main">הורדת גיבוי</span></button>
        <label class="row action">${icon('tray.and.arrow.down')}<span class="row-main">שחזור מגיבוי</span><input type="file" id="file-backup" accept=".json,application/json" hidden></label>
        <button class="row action destructive" data-act="reset" ${demo ? 'disabled' : ''}>${icon('trash')}<span class="row-main">מחיקת כל הנתונים</span></button>
      </div>
      <p class="list-footer">${store?.kind === 'db' ? 'הנתונים נשמרים בחשבון שלכם ב-claude.ai, ורק אתם רואים אותם.' : 'הנתונים נשמרים בדפדפן הזה בלבד. כדאי להוריד גיבוי מדי פעם, ולהעביר איתו את העבודה למחשב אחר.'}</p>
    </section>`;
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
  notices.roster = { busy: true, text: 'קורא את הקובץ…' };
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
  notices.schedule = { busy: true, text: 'קורא את הלו"ז… (כמה שניות)' };
  render();
  try {
    const pdfjs = await loadPdfjs();
    const { cells, warnings } = await extractScheduleCells(pdfjs, new Uint8Array(await file.arrayBuffer()));
    if (!cells.length) throw new Error('לא נמצאו תאים עם תאריכים. האם זה קובץ הלו"ז השבועי?');
    leaveDemo();
    const parsed = parseCells(cells, createNameMatcher(people().map(p => p.name)));
    const firstImport = !state.events.some(e => !e.manual);
    const result = mergeSchedule(state.events, parsed);
    state.events = result.events;
    const dates = parsed.map(e => e.date).sort();
    state.imports.schedule = [...(state.imports.schedule || []).slice(-9), { name: file.name, at: new Date().toISOString(), events: parsed.length, from: dates[0], to: dates[dates.length - 1] }];
    const relevant = parsed.filter(e => isRelevant(e)).length;
    let text = `נמצאו ${parsed.length} תאים, מהם ${relevant} ימי עיון וקורסים עם משבצות לשיבוץ. חדשים: ${result.added}, עודכנו: ${result.updated}${result.removed ? `, הוסרו: ${result.removed}` : ''}.`;
    if (result.kept.length) text += ` ${result.kept.length} אירועים ששובצו בעבר לא מופיעים בקובץ החדש ומסומנים בלוח.`;
    if (warnings.length) text += ' ' + warnings.join(' ');
    // On the first import everything is new; the report starts with the next update.
    state.lastImport = { name: file.name, at: new Date().toISOString(), report: result.report, dismissed: firstImport };
    const attention = relevantChanges().filter(r => r.level === 'attention').length;
    if (attention) text += ` ${attention} שינויים דורשים בדיקה - מפורטים בלשונית "שיבוץ".`;
    notices.schedule = { text };
    if (view === 'board' && result.report.length) toast(attention ? `הלו"ז עודכן · ${attention} שינויים לבדיקה` : 'הלו"ז עודכן');
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

function renderChrome() {
  $('#nav-title').textContent = VIEW_TITLES[view];
  for (const v of Object.keys(VIEW_TITLES)) {
    $(`#view-${v}`).hidden = v !== view;
    const tab = $(`#tab-${v}`);
    if (v === view) tab.setAttribute('aria-current', 'page');
    else tab.removeAttribute('aria-current');
  }
  // Bar items: bare symbols; the screen's main action stays in the content.
  const actions = {
    board: hasData(state) ? `
      <label class="icon-btn" title='עדכון לו"ז' aria-label='עדכון לו"ז'>${icon('arrow.clockwise')}<input type="file" id="file-schedule-quick" accept=".pdf,application/pdf" hidden></label>
      <button class="icon-btn" data-act="add-event" title="אירוע ידני" aria-label="אירוע ידני">${icon('plus')}</button>` : '',
    people: `<button class="icon-btn" data-act="add-person" title="הוספת מדריך/ה" aria-label="הוספת מדריך/ה">${icon('person.badge.plus')}</button>`,
    export: `<button class="icon-btn" data-act="copy-doc" title="העתקת הטקסט" aria-label="העתקת הטקסט">${icon('doc.on.doc')}</button>`,
    files: '',
  };
  $('#nav-actions').innerHTML = actions[view];
}

function render() {
  renderChrome();
  if (view === 'board') renderBoard();
  else if (view === 'people') renderPeople();
  else if (view === 'export') renderExport();
  else renderFiles();
}

function go(v) {
  view = v;
  closeMenu();
  render();
  window.scrollTo({ top: 0 });
}

/** Scroll to an event's card, widening the filters if they hide it. */
function jumpTo(evId) {
  const find = () => document.querySelector(`[data-ev="${CSS.escape(evId)}"]`);
  let card = find();
  if (!card) {
    const ev = findEvent(evId);
    if (!ev) return;
    const s = state.settings;
    s.month = 'all';
    s.onlyOpen = false;
    if (ev.date < todayISO()) s.showPast = true;
    commit();
    card = find();
  }
  if (!card) { toast('האירוע מוסתר בסינון הנוכחי'); return; }
  card.scrollIntoView({ behavior: 'smooth', block: 'center' });
  card.classList.add('flash');
  setTimeout(() => card.classList.remove('flash'), 1600);
}

// ---------------------------------------------------------------- events

document.addEventListener('click', async e => {
  const t = e.target;

  // Menus close on any click; a click on an item runs it.
  if (menuHandlers) {
    const item = t.closest('[data-menu]');
    const handlers = menuHandlers;
    closeMenu();
    if (item) { handlers[+item.dataset.menu].run(); return; }
    if (t.closest('[data-act="ev-menu"]')) return;
  }
  const alertBtn = t.closest('[data-alert]');
  if (alertBtn) { const a = alertActions[+alertBtn.dataset.alert]; $('#alert').close(); a.run?.(); return; }

  const jump = t.closest('[data-jump]');
  if (jump) { jumpTo(jump.dataset.jump); return; }
  const goBtn = t.closest('[data-go]');
  if (goBtn) { go(goBtn.dataset.go); return; }
  const tab = t.closest('.tab');
  if (tab) { go(tab.dataset.view); return; }
  if (t.closest('[data-close]')) { closeDialog(); return; }

  // Sheet actions
  if (t.closest('#dialog')) {
    const pick = t.closest('[data-pick]');
    if (pick && dialogHandlers.pick) { dialogHandlers.pick(pick.dataset.pick); return; }
    const gact = t.closest('[data-gact]');
    if (gact && dialogHandlers.group) {
      e.preventDefault();
      const row = gact.closest('[data-gi]');
      dialogHandlers.group(gact.dataset.gact, row ? +row.dataset.gi : -1);
      return;
    }
    const act = t.closest('[data-act]')?.dataset.act;
    if (act === 'pick-free') { const v = $('#pick-free').value.trim(); if (v) dialogHandlers.pick(v); return; }
    if (act === 'pick-clear') { dialogHandlers.clear(); return; }
    if (act === 'sheet-done') { dialogHandlers.save?.(); return; }
    if (act === 'slot-delete') { dialogHandlers.remove(); return; }
    if (act === 'avail-reset') { dialogHandlers.reset(); return; }
    if (act === 'remind') {
      const p = people().find(x => x.id === t.closest('[data-person]').dataset.person);
      if (p) copyText(reminderText(p));
      return;
    }
    if (t === dialog()) closeDialog(); // backdrop
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
    case 'ev-menu': eventMenu(actEl, ev); break;
    case 'dismiss-changes':
      if (state.lastImport) state.lastImport.dismissed = true;
      state.settings.onlyChanged = false;
      commit();
      break;
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
      if (n) { autoUndo = snapshot; toast(`שובצו ${n} עמדות - הכחולות שובצו אוטומטית, כדאי לעבור עליהן`); }
      else toast('לא נמצאו מדריכים פנויים לעמדות החסרות');
      commit();
      break;
    }
    case 'undo-auto':
      if (autoUndo) { state.events = autoUndo; autoUndo = null; toast('השיבוץ האוטומטי בוטל'); commit(); }
      break;
    case 'add-person': personEditor(); break;
    case 'open-person': {
      const p = people().find(x => x.id === actEl.dataset.person);
      if (p) personSheet(p);
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
      try { await saveFile(`גיבוי שיבוץ מדריכים ${todayISO()}.json`, JSON.stringify(state)); } catch (err) { toast('ההורדה נכשלה: ' + (err.message || err.code)); }
      break;
    }
    case 'reset':
      openAlert({
        title: 'למחוק את כל הנתונים?',
        message: 'המדריכים, האירועים, השיבוצים וההערות יימחקו מהאתר. אי אפשר לבטל את זה - כדאי להוריד גיבוי קודם.',
        actions: [
          { label: 'ביטול', cls: 'btn-gray' },
          { label: 'מחיקה', cls: 'btn-prominent btn-destructive', run: () => {
            state = emptyState();
            autoUndo = null;
            notices.roster = notices.schedule = null;
            toast('כל הנתונים נמחקו');
            commit();
          } },
        ],
      });
      break;
    default: break;
  }
});

document.addEventListener('change', e => {
  const t = e.target;
  if (t.closest('#dialog')) return; // sheets save on Done
  const s = state.settings;
  if (t.dataset.filter) { s[t.dataset.filter][t.value] = t.checked; commit(); return; }
  if (t.dataset.toggle) { s[t.dataset.toggle] = t.checked; commit(); return; }
  if (t.dataset.setting === 'month') { s.month = t.value; commit(); return; }

  if (t.id === 'x-month') { state.exportOptions.month = t.value; commit(); return; }
  if (t.id === 'x-location') { state.exportOptions.location = t.checked; commit(); return; }
  if (t.dataset.section) { state.exportOptions.sections[t.dataset.section] = t.checked; commit(); return; }

  if (t.id === 'file-roster' && t.files[0]) importRoster(t.files[0]);
  if ((t.id === 'file-schedule' || t.id === 'file-schedule-quick') && t.files[0]) importSchedule(t.files[0]);
  if (t.id === 'file-backup' && t.files[0]) restoreBackup(t.files[0]);
  // The schedule file often keeps its name; clearing lets the same file be picked again.
  if (t.type === 'file') setTimeout(() => { t.value = ''; }, 0);
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
  if (e.key === 'Escape' && menuHandlers) { closeMenu(); return; }
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
document.addEventListener('toggle', e => { if (e.target.id === 'filters-box') filtersOpen = e.target.open; }, true);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });
window.addEventListener('resize', closeMenu);
// The bar firms into a band once content scrolls beneath it.
const navbar = $('#navbar');
const onScroll = () => {
  navbar.dataset.scrolled = String(window.scrollY > 24);
  if (menuHandlers && Math.abs(window.scrollY - menuScrollY) > 40) closeMenu();
};
window.addEventListener('scroll', onScroll, { passive: true });

// ---------------------------------------------------------------- start

async function start() {
  document.querySelectorAll('[data-icon]').forEach(el => { el.outerHTML = icon(el.dataset.icon); });
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
  }
  personIndex = null;
  render();
}

start();
