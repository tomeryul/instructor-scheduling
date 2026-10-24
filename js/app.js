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
import { stampChanges, mergeStates, describeChanges, packState, unpackState, sharedState } from './sync.js';
import { parseRepo, checkRepo, readFile, writeFile, deleteFile, connectLink, readConnectLink, DEFAULT_PATH } from './github-sync.js';
import { demoState } from './demo.js';
import { icon } from './icons.js';

const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const DOCX_URL = 'https://cdn.jsdelivr.net/npm/docx@9.0.2/build/index.umd.js';
const QR_URL = 'https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js';
const DEVICE_KEY = 'nsm-shibutz-device';
const SYNC_KEY = 'nsm-shibutz-sync';

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
    deleted: { events: {}, slots: {} },
    updatedAt: 0,
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
    deleted: { events: { ...(s.deleted?.events || {}) }, slots: { ...(s.deleted?.slots || {}) } },
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
let snapshot = null; // the state as of the last commit, to see what changed

/** Stamp what changed since the last commit, save it and hand it to sync - without redrawing. */
function persist() {
  if (demo) return;
  if (stampChanges(snapshot, state)) state.updatedAt = Date.now();
  snapshot = clone(state);
  dirty = true;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flush, 500);
  noteLocalChange();
}

function commit() {
  personIndex = null;
  render();
  persist();
}

/** Remember a deletion, so a merge with another device does not bring the item back. */
function markDeleted(kind, id) {
  state.deleted[kind][id] = Date.now();
  state.updatedAt = Date.now();
}

/** Swap in a state that came from another device: it carries its own stamps. */
function replaceState(next) {
  state = withDefaults(next);
  snapshot = clone(state);
  syncState.localHash = sharedHash(state);
  personIndex = null;
  autoUndo = null;
  render();
  dirty = true;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flush, 300);
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
  snapshot = clone(state);
  autoUndo = null;
}

// ---------------------------------------------------------------- feedback

let toastTimer = null;
/** A short message; `action` ({label, run}) adds one button, e.g. Undo. */
function toast(msg, action) {
  const el = $('#toast');
  el.innerHTML = `<span>${esc(msg)}</span>${action ? `<button type="button" class="toast-action">${esc(action.label)}</button>` : ''}`;
  if (action) el.querySelector('button').onclick = () => { el.hidden = true; action.run(); };
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, action ? 8000 : 3400);
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
  m.innerHTML = items.map((it, i) => {
    if (it === '-') return '<hr>';
    if (it.header) return `<p class="menu-header">${esc(it.header)}</p>`;
    const label = it.sub ? `<span class="menu-label">${esc(it.label)}<small>${esc(it.sub)}</small></span>` : esc(it.label);
    return `<button role="menuitem" data-menu="${i}" class="${it.destructive ? 'destructive' : ''}" ${it.disabled ? 'disabled' : ''}>${label}${icon(it.icon)}</button>`;
  }).join('');
  m.hidden = false;
  const r = anchor.getBoundingClientRect();
  const w = m.offsetWidth, h = m.offsetHeight;
  let top = r.bottom + 6;
  if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 6);
  const left = Math.min(Math.max(8, r.left), window.innerWidth - w - 8);
  m.style.top = `${top}px`;
  m.style.left = `${left}px`;
  menuScrollY = window.scrollY;
  m.querySelector('button:not(:disabled)')?.focus({ preventScroll: true });
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
        <span class="row-sub">השמות והאירועים כאן מומצאים, כדי להראות איך זה עובד. שינויים בדוגמה לא נשמרים. כבר עובדים במכשיר אחר? בלשונית "קבצים" מביאים את הנתונים משם.</span></div>
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
    ${syncNotice()}
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
      markDeleted('slots', slot.id);
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
    ...(ev.missingFromSchedule || ev.manual ? [{ label: 'מחיקה', icon: 'trash', destructive: true, run: () => { state.events = state.events.filter(x => x !== ev); markDeleted('events', ev.id); commit(); toast('האירוע נמחק'); } }] : []),
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
    ${devicesSection()}
    <section class="list-section">
      <h3 class="list-header">העברה בקובץ</h3>
      <div class="list">
        <button class="row action" data-act="send" ${demo ? 'disabled' : ''}>${icon('square.and.arrow.up')}<span class="row-main">שליחה למכשיר אחר</span></button>
        <label class="row action">${icon('tray.and.arrow.down')}<span class="row-main">קבלה ממכשיר אחר</span><input type="file" id="file-receive" accept=".json,.txt,application/json,text/plain" hidden></label>
      </div>
      <p class="list-footer">${transferFooter()}</p>
    </section>
    <section class="list-section">
      <div class="list">
        ${syncOn() ? `<button class="row action destructive" data-act="sync-delete">${icon('trash')}<span class="row-main">מחיקת הנתונים מ-GitHub</span></button>` : ''}
        <button class="row action destructive" data-act="reset" ${demo ? 'disabled' : ''}>${icon('trash')}<span class="row-main">מחיקת כל הנתונים מהמכשיר הזה</span></button>
      </div>
    </section>`;
}

function devicesSection() {
  if (store?.kind === 'db') {
    return `
    <section class="list-section">
      <h3 class="list-header">בין מכשירים</h3>
      <div class="list"><div class="row"><span class="row-icon ok">${icon('checkmark')}</span><span class="row-main"><span class="row-title">מסונכרן דרך החשבון ב-claude.ai</span><span class="row-sub">פותחים את אותו קישור בכל מכשיר שמחובר לחשבון, והנתונים שם. רק אתם רואים אותם.</span></span></div></div>
    </section>`;
  }
  if (syncCfg) {
    return `
    <section class="list-section">
      <h3 class="list-header">סנכרון עם GitHub</h3>
      <div class="list">
        <div class="row" id="sync-status">${syncStatusHtml()}</div>
        <button class="row action" data-act="sync-push" ${demo ? 'disabled' : ''}>${icon('square.and.arrow.up')}<span class="row-main"><span class="row-title">שמירה ל-GitHub</span><span class="row-sub">בסוף עבודה: שולח את מה שנעשה במכשיר הזה</span></span></button>
        <button class="row action" data-act="sync-pull">${icon('square.and.arrow.down')}<span class="row-main"><span class="row-title">ייבוא מ-GitHub</span><span class="row-sub">בתחילת עבודה: מביא את מה שנשמר במכשירים אחרים</span></span></button>
        <button class="row action" data-act="sync-both">${icon('arrow.triangle.2.circlepath')}<span class="row-main"><span class="row-title">סנכרון</span><span class="row-sub">ייבוא ושמירה בלחיצה אחת</span></span></button>
        <label class="row"><span class="row-main"><span class="row-title">סנכרון אוטומטי</span><span class="row-sub">שומר ומייבא לבד, בלי ללחוץ</span></span><input type="checkbox" role="switch" class="switch" id="sync-auto" ${syncCfg.auto ? 'checked' : ''}></label>
      </div>
      <p class="list-footer">כל שינוי נשמר במכשיר הזה מיד, גם בלי ללחוץ על כלום. "שמירה" ו"ייבוא" מעבירים עבודה בין המכשיר הזה ל-GitHub, והם זמינים בכל מסך בכפתור החצים שבראש המסך. שום דבר לא נדרס: בכל עמדה נשאר השיבוץ האחרון שנעשה, באיזה מכשיר שלא יהיה.</p>
    </section>
    <section class="list-section">
      <div class="list">
        <button class="row action" data-act="sync-share">${icon('qrcode')}<span class="row-main">חיבור מכשיר נוסף</span>${chevron()}</button>
        <button class="row action" data-act="sync-off">${icon('xmark.circle.fill')}<span class="row-main">ניתוק המכשיר הזה</span></button>
      </div>
    </section>`;
  }
  const iosTip = isIOS() && !isStandalone()
    ? ' באייפון כדאי גם להוסיף את האתר למסך הבית (שיתוף ← הוספה למסך הבית): Safari עלול למחוק נתונים של אתר שלא נפתח כמה ימים. לאתר במסך הבית יש אחסון נפרד, אז מעבירים אליו את הנתונים פעם אחת.'
    : '';
  return `
    <section class="list-section">
      <h3 class="list-header">בין מכשירים</h3>
      <div class="list">
        <button class="row" data-act="sync-setup"><span class="row-icon sync">${icon('arrow.triangle.2.circlepath')}</span><span class="row-main"><span class="row-title">סנכרון אוטומטי דרך GitHub</span><span class="row-sub">הגדרה של פעם אחת בכל מכשיר, ומאז הכול מתעדכן לבד</span></span>${chevron()}</button>
      </div>
      <p class="list-footer">בלי סנכרון, כל מכשיר שומר עותק משלו ומעבירים עבודה בקובץ.${iosTip}</p>
    </section>`;
}

function transferFooter() {
  const how = isTouch()
    ? 'נפתח תפריט השיתוף: WhatsApp לעצמכם, מייל, AirDrop או שמירה בקבצים.'
    : 'הקובץ יורד למחשב; שולחים אותו לעצמכם (מייל, WhatsApp, Drive).';
  const log = [
    device.sent ? `נשלח מכאן לאחרונה ${whenLabel(device.sent)}.` : '',
    device.received ? `התקבל לאחרונה מ-${esc(isolate(device.received.from))} ${whenLabel(device.received.at)}.` : '',
  ].filter(Boolean).join(' ');
  return `${how} במכשיר השני בוחרים "קבלה ממכשיר אחר": הנתונים מתמזגים עם מה שיש שם, ובכל עמדה נשאר השיבוץ האחרון שנעשה. ${log}`;
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

// ---------------------------------------------------------------- devices

const readLocal = key => { try { return JSON.parse(localStorage.getItem(key) || 'null'); } catch { return null; } };
const writeLocal = (key, value) => {
  try { if (value == null) localStorage.removeItem(key); else localStorage.setItem(key, JSON.stringify(value)); } catch { /* storage is off */ }
};

function guessDeviceName() {
  const ua = navigator.userAgent;
  if (/iPhone/.test(ua)) return 'iPhone';
  if (/iPad/.test(ua) || (/Macintosh/.test(ua) && navigator.maxTouchPoints > 1)) return 'iPad';
  if (/Android/.test(ua)) return /Mobile/.test(ua) ? 'טלפון Android' : 'טאבלט Android';
  if (/Macintosh|Mac OS X/.test(ua)) return 'Mac';
  if (/Windows/.test(ua)) return 'מחשב Windows';
  if (/CrOS/.test(ua)) return 'Chromebook';
  if (/Linux/.test(ua)) return 'מחשב Linux';
  return 'מכשיר אחר';
}

/** This device: its name, and when work last left it or arrived. */
const device = { name: guessDeviceName(), sent: null, received: null, ...(readLocal(DEVICE_KEY) || {}) };
const saveDevice = () => writeLocal(DEVICE_KEY, device);
const isTouch = () => window.matchMedia('(pointer: coarse)').matches;
const isIOS = () => /iPhone|iPad/.test(navigator.userAgent) || (/Macintosh/.test(navigator.userAgent) && navigator.maxTouchPoints > 1);
const isStandalone = () => window.matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;

function whenLabel(t) {
  const d = new Date(t);
  if (!t || Number.isNaN(d.getTime())) return '';
  const ago = Date.now() - d.getTime();
  if (ago < 60000) return 'לפני רגע';
  if (ago < 3600000) return `לפני ${Math.round(ago / 60000)} דק׳`;
  const hm = d.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' });
  if (d.toDateString() === new Date().toDateString()) return `היום ${hm}`;
  if (d.toDateString() === new Date(Date.now() - 86400000).toDateString()) return `אתמול ${hm}`;
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')} ${hm}`;
}

// Device and repository names are often Latin: isolate them so the Hebrew sentence keeps its order.
const isolate = text => `\u2068${text}\u2069`;

/** "התקבל מ-iPhone: 5 שיבוצים, לו"ז עדכני יותר", or '' when nothing changed. */
function changesText(d, from) {
  const n = (k, one, many) => (k === 1 ? one : `${k} ${many}`);
  const parts = [];
  if (d.assigned) parts.push(n(d.assigned, 'שיבוץ אחד', 'שיבוצים'));
  if (d.cleared) parts.push(n(d.cleared, 'עמדה אחת שפונתה', 'עמדות שפונו'));
  if (d.events) parts.push(n(d.events, 'אירוע אחד', 'אירועים'));
  if (d.schedule) parts.push('לו"ז עדכני יותר');
  if (d.roster) parts.push('רשימת מדריכים עדכנית יותר');
  if (d.people) parts.push(n(d.people, 'עדכון במדריך/ה', 'עדכונים במדריכים'));
  if (d.notes) parts.push('הערות לייצוא');
  return parts.length ? `התקבל מ-${isolate(from)}: ${parts.join(', ')}` : '';
}

let persistAsked = false;
/** Ask the browser to keep this site's data, once there is data worth keeping. */
async function askToPersist() {
  if (persistAsked || store?.kind !== 'local' || !navigator.storage?.persist) return;
  persistAsked = true;
  try { if (!(await navigator.storage.persisted())) await navigator.storage.persist(); } catch { /* not offered here */ }
}

// ---------------------------------------------------------------- moving work between devices in a file

async function sendToDevice() {
  const hash = syncState.localHash;
  const text = JSON.stringify(packState(state, device.name));
  const filename = `שיבוץ מדריכים ${todayISO()} ${device.name}.json`;
  const sent = () => {
    Object.assign(device, { sent: new Date().toISOString(), sentHash: hash });
    saveDevice();
    renderSyncBits();
    if (view === 'files') renderFiles();
  };
  // Phones: the share sheet (WhatsApp to yourself, mail, AirDrop, Files). Some browsers only share .txt.
  if (isTouch() && navigator.canShare) {
    for (const [ext, type] of [['json', 'application/json'], ['txt', 'text/plain']]) {
      const file = new File([text], filename.replace(/json$/, ext), { type });
      if (!navigator.canShare({ files: [file] })) continue;
      try { await navigator.share({ files: [file], title: 'שיבוץ מדריכים' }); sent(); return; } catch (e) { if (e.name === 'AbortError') return; break; }
    }
  }
  try {
    if (await saveFile(filename, text) === 'saved') { sent(); toast('הקובץ ירד. במכשיר השני: קבלה ממכשיר אחר'); }
  } catch (e) { toast('ההורדה נכשלה: ' + (e.message || e.code)); }
}

/** Choose a file sent from another device (from a menu, where there is no file field). */
function pickReceiveFile() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json,.txt,application/json,text/plain';
  input.addEventListener('change', () => { if (input.files[0]) receiveFile(input.files[0]); });
  input.click();
}

async function receiveFile(file) {
  let data = null;
  try { data = JSON.parse(await file.text()); } catch { data = null; }
  const incoming = unpackState(data);
  if (!incoming) { toast('זה לא קובץ שנשלח מהאתר'); return; }
  const from = data.device || 'הקובץ';
  const wasDemo = demo;
  const wasClean = !wasDemo && syncState.localHash === device.sentHash;
  const sentHashBefore = device.sentHash;
  leaveDemo();
  const before = clone(state);
  const merged = mergeStates(state, incoming);
  const text = changesText(describeChanges(before, merged), from);
  replaceState(merged);
  // Nothing was waiting to be sent, and what arrived is already on the other device: still nothing to send.
  if (wasClean || wasDemo) device.sentHash = syncState.localHash;
  device.received = { at: new Date().toISOString(), from };
  saveDevice();
  renderSyncBits();
  askToPersist();
  if (!text) { toast(`אין חדש מ-${isolate(from)} - המכשיר הזה כבר מעודכן`); return; }
  toast(text, wasDemo ? null : {
    label: 'ביטול',
    run: () => { replaceState(before); device.sentHash = sentHashBefore; saveDevice(); renderSyncBits(); toast('הקבלה בוטלה'); },
  });
}

// ---------------------------------------------------------------- sync through GitHub: save, import, or both

// syncCfg: { owner, repo, path, token, auto,
//   sha, etag, remoteHash, remoteDevice, remoteAt - the file as last seen, who wrote it and when;
//   importedSha - the version of the file last brought into this device;
//   cleanHash - this device's data as it was last saved (or when it had nothing unsaved);
//   savedAt, importedAt, at }
let syncCfg = readLocal(SYNC_KEY);
const syncState = { busy: null, pending: null, queued: null, error: null, timer: null, autoPushes: 0, localHash: null };
const syncOn = () => !!syncCfg && store?.kind === 'local';
const syncAuto = () => syncOn() && !!syncCfg.auto;
const saveSyncCfg = () => writeLocal(SYNC_KEY, syncCfg);
const BUSY_TEXT = { push: 'שומר ב-GitHub…', pull: 'מייבא מ-GitHub…', both: 'מסנכרן…', check: 'בודק אם יש עדכון…', delete: 'מוחק מ-GitHub…' };

// Two requests combined: the same, the one that does more, or both directions.
const strongest = (a, b) => (!a ? b : !b ? a : a === b ? a : a === 'check' ? b : b === 'check' ? a : 'both');

// Never swap the data under an open sheet, menu or a field being typed in.
const uiBusy = () => $('#dialog').open || $('#alert').open || !!menuHandlers
  || !!document.activeElement?.matches?.('main textarea, main input[type="search"], main input[type="text"], main input:not([type])');

/** JSON with sorted keys, so equal states compare equal whatever order they were built in. */
const canon = x => JSON.stringify(x, (k, v) => (v && typeof v === 'object' && !Array.isArray(v) ? Object.fromEntries(Object.keys(v).sort().map(key => [key, v[key]])) : v));
function hashText(str) {
  let a = 0x811c9dc5, b = 0x9747b28c;
  for (let i = 0; i < str.length; i++) { const c = str.charCodeAt(i); a = Math.imul(a ^ c, 16777619); b = Math.imul(b ^ c, 2246822507); }
  return (a >>> 0).toString(36) + (b >>> 0).toString(36);
}
const sharedHash = s => hashText(canon(sharedState(s)));

/** Called after every change on this device. */
function noteLocalChange() {
  syncState.localHash = sharedHash(state);
  syncState.autoPushes = 0;
  renderSyncBits();
  if (syncAuto()) scheduleSync('both');
}

/** What is waiting: work on this device that was not saved, or work in the file that was not imported here. */
function syncFlags() {
  if (store?.kind !== 'local' || demo || !hasData(state)) return { unsaved: false, news: false };
  if (syncOn()) return { unsaved: syncState.localHash !== syncCfg.cleanHash, news: !!syncCfg.sha && syncCfg.sha !== syncCfg.importedSha };
  // Without GitHub, only once files have gone between devices - one device on its own is never "unsaved".
  const usesFiles = !!(device.sentHash || device.received);
  return { unsaved: usesFiles && syncState.localHash !== device.sentHash, news: false };
}

function scheduleSync(mode = 'both', delay = 2500) {
  if (!syncOn()) return;
  syncState.queued = strongest(syncState.queued, mode);
  clearTimeout(syncState.timer);
  syncState.timer = setTimeout(() => { const m = syncState.queued; syncState.queued = null; runSync(m); }, delay);
}

function resumePending(delay = 800) {
  if (!syncState.pending || syncState.busy) return;
  const m = syncState.pending;
  syncState.pending = null;
  scheduleSync(m, delay);
}

/**
 * push  - save: write this device's work into the file (merged with anything
 *         there this device has not imported, so nothing is lost);
 * pull  - import: merge the file into this device;
 * both  - sync: import, then save;
 * check - only look whether the file changed, for the "update waiting" mark.
 */
async function runSync(mode = 'both', { manual = false } = {}) {
  if (!syncOn()) return;
  if (syncState.busy) { syncState.pending = strongest(syncState.pending, mode); return; }
  const cfg = syncCfg;
  syncState.busy = mode;
  renderSyncBits();
  let imported = '';
  let wrote = false;
  let leftUnimported = false;
  try {
    for (let attempt = 0; ; attempt++) {
      if (attempt > 6) throw new Error('GitHub לא הגיב כמצופה - נסו שוב');
      const remote = await readFile(cfg, cfg.sha && cfg.remoteHash ? cfg.etag : null);
      let incoming = null;
      if (remote && !remote.notModified) {
        incoming = unpackState(remote.data);
        if (!incoming) throw new Error('הקובץ ב-GitHub לא נראה כמו קובץ של האתר');
        Object.assign(cfg, { sha: remote.sha, etag: remote.etag, remoteHash: sharedHash(incoming), remoteDevice: remote.data.device || null, remoteAt: remote.data.savedAt || null });
      } else if (!remote) {
        if (cfg.sha) cfg.cleanHash = null;
        Object.assign(cfg, { sha: null, etag: null, remoteHash: null, remoteDevice: null, remoteAt: null });
      }
      const news = !!cfg.sha && cfg.sha !== cfg.importedSha;
      if (mode === 'check') break;
      if (news && !incoming) { cfg.etag = null; continue; } // the file changed: read it in full

      if ((mode === 'pull' || mode === 'both') && news) {
        if (uiBusy()) { syncState.pending = strongest(syncState.pending, mode); return; }
        const wasDemo = demo && hasData(withDefaults(incoming));
        if (wasDemo) leaveDemo();
        if (!demo) {
          const wasClean = wasDemo || syncState.localHash === cfg.cleanHash;
          const before = clone(state);
          const merged = mergeStates(state, incoming);
          if (canon(sharedState(merged)) !== canon(sharedState(before))) {
            replaceState(merged);
            imported = wasDemo ? `הנתונים נטענו מ-${isolate(`${cfg.owner}/${cfg.repo}`)}` : changesText(describeChanges(before, merged), cfg.remoteDevice || 'מכשיר אחר');
          }
          if (wasClean) cfg.cleanHash = syncState.localHash;
          cfg.importedAt = new Date().toISOString();
        }
        cfg.importedSha = cfg.sha;
      }

      if (mode === 'push' || mode === 'both') {
        if (demo) break; // never write the sample data
        const unimported = !!cfg.sha && cfg.sha !== cfg.importedSha;
        const out = unimported ? mergeStates(state, incoming) : state;
        const outHash = unimported ? sharedHash(out) : syncState.localHash;
        // Automatic runs stop after a few writes without a new edit, so two devices cannot bounce a change.
        if (outHash !== cfg.remoteHash && (manual || syncState.autoPushes < 4)) {
          if (!manual) syncState.autoPushes++;
          try {
            const sha = await writeFile(cfg, JSON.stringify(packState(out, device.name)), cfg.sha, `שמירה מ-${device.name}`);
            Object.assign(cfg, { sha, etag: null, remoteHash: outHash });
            if (!unimported) Object.assign(cfg, { importedSha: sha, remoteDevice: device.name, remoteAt: new Date().toISOString() });
            wrote = true;
          } catch (e) {
            if (e.code === 'conflict') { cfg.etag = null; continue; }
            throw e;
          }
        }
        leftUnimported = unimported;
        cfg.cleanHash = syncState.localHash;
        cfg.savedAt = new Date().toISOString();
      }
      break;
    }
    cfg.at = new Date().toISOString();
    syncState.error = null;
    const waiting = leftUnimported ? ` · יש שם גם עדכון מ-${isolate(cfg.remoteDevice || 'מכשיר אחר')} שעוד לא יובא לכאן` : '';
    if (mode === 'pull' && manual) toast(imported || (cfg.sha ? 'אין חדש ב-GitHub - המכשיר הזה כבר מעודכן' : 'ב-GitHub אין כרגע נתונים לייבוא'));
    else if (mode === 'push' && manual) toast((wrote ? 'נשמר ב-GitHub' : 'כבר שמור ב-GitHub - אין שינויים חדשים') + waiting);
    else if (mode === 'both' && manual) toast([imported, wrote ? 'נשמר ב-GitHub' : ''].filter(Boolean).join(' · ') || 'הכול מסונכרן - אין שינויים');
    else if (imported) toast(imported);
  } catch (e) {
    syncState.error = e;
    if (manual) toast(e.message);
  } finally {
    if (syncCfg === cfg) saveSyncCfg();
    syncState.busy = null;
    renderSyncBits();
    if (!uiBusy()) resumePending();
  }
}

/** One line on where this device stands. */
function syncSummary() {
  if (syncState.busy) return BUSY_TEXT[syncState.busy];
  const e = syncState.error;
  if (e && e.code === 'offline') return 'אין חיבור לאינטרנט. העבודה שמורה במכשיר הזה.';
  if (e) return `הפעולה האחרונה נכשלה: ${e.message}`;
  const f = syncFlags();
  if (!syncCfg.sha) return demo || !hasData(state) ? 'ב-GitHub אין כרגע נתונים.' : 'ב-GitHub אין כרגע נתונים. "שמירה" תעלה לשם את מה שיש במכשיר הזה.';
  const lines = [f.unsaved ? 'יש במכשיר הזה שינויים שלא נשמרו ב-GitHub.' : 'כל העבודה במכשיר הזה שמורה ב-GitHub.'];
  if (f.news) lines.push(`יש עדכון מ-${isolate(syncCfg.remoteDevice || 'מכשיר אחר')}${syncCfg.remoteAt ? ` (${whenLabel(syncCfg.remoteAt)})` : ''} שעוד לא יובא.`);
  return lines.join(' ');
}

function fileSummary() {
  const f = syncFlags();
  if (f.unsaved) return 'יש שינויים מאז השליחה האחרונה למכשיר אחר.';
  if (device.sent) return `נשלח לאחרונה ${whenLabel(device.sent)}.`;
  return 'כל מכשיר שומר עותק משלו. כך מעבירים עבודה ביניהם:';
}

function syncButtonHtml() {
  if (store?.kind !== 'local') return '';
  const f = syncFlags();
  const failed = syncOn() && syncState.error && syncState.error.code !== 'offline';
  const parts = [syncOn() ? 'שמירה, ייבוא וסנכרון' : 'שמירה וייבוא'];
  if (f.unsaved) parts.push('יש שינויים שלא נשמרו');
  if (f.news) parts.push('יש עדכון לייבוא');
  if (failed) parts.push('הפעולה האחרונה נכשלה');
  const label = parts.join(' · ');
  return `<button class="icon-btn${syncState.busy ? ' busy' : ''}" data-act="sync-menu" aria-label="${esc(label)}" title="${esc(label)}">${icon('arrow.triangle.2.circlepath')}${f.unsaved || f.news || failed ? '<span class="badge-dot" aria-hidden="true"></span>' : ''}</button>`;
}

function syncMenu(anchor) {
  const items = syncOn()
    ? [
      { header: syncSummary() },
      { label: 'שמירה ל-GitHub', sub: 'שולח את העבודה מהמכשיר הזה', icon: 'square.and.arrow.up', disabled: demo, run: () => runSync('push', { manual: true }) },
      { label: 'ייבוא מ-GitHub', sub: 'מביא את מה שנשמר במכשירים אחרים', icon: 'square.and.arrow.down', run: () => runSync('pull', { manual: true }) },
      { label: 'סנכרון', sub: 'ייבוא ושמירה בלחיצה אחת', icon: 'arrow.triangle.2.circlepath', run: () => runSync('both', { manual: true }) },
      '-',
      { label: 'הגדרות סנכרון', icon: 'slider', run: () => go('files') },
    ]
    : [
      { header: fileSummary() },
      { label: 'שליחה למכשיר אחר', sub: 'קובץ עם כל העבודה', icon: 'square.and.arrow.up', disabled: demo, run: sendToDevice },
      { label: 'קבלה ממכשיר אחר', sub: 'ממזג קובץ שנשלח ממכשיר אחר', icon: 'square.and.arrow.down', run: pickReceiveFile },
      '-',
      { label: 'סנכרון דרך GitHub', icon: 'arrow.triangle.2.circlepath', run: syncSetup },
    ];
  openMenu(anchor, items);
}

/** Refresh every place that shows the sync state, without redrawing the page under the user. */
let shownSyncKey = '';
function renderSyncBits() {
  const nav = $('#nav-leading');
  if (nav) nav.innerHTML = syncButtonHtml();
  const el = $('#sync-status');
  if (el) el.innerHTML = syncStatusHtml();
  // The board's notices (a failure, an update to import) change only when the page is free.
  const e = syncState.error;
  const key = `${syncOn() && e && e.code !== 'offline' ? e.message : ''}|${syncFlags().news ? syncCfg.sha : ''}`;
  if (key !== shownSyncKey && !uiBusy()) {
    shownSyncKey = key;
    if (view === 'board') renderBoard();
  }
}

function syncStatusHtml() {
  if (!syncCfg) return '';
  const e = syncState.error;
  const bad = e && e.code !== 'offline';
  return `<span class="row-icon ${bad ? 'bad' : 'sync'}">${icon(bad ? 'exclamation.triangle.fill' : 'arrow.triangle.2.circlepath')}</span>
    <span class="row-main"><span class="row-title"><bdi>${esc(`${syncCfg.owner}/${syncCfg.repo}`)}</bdi></span><span class="row-sub">${esc(syncSummary())}</span></span>`;
}

function syncNotice() {
  if (!syncOn()) return '';
  const e = syncState.error;
  if (e && e.code !== 'offline') {
    return `
    <div class="list"><div class="notice">
      <span class="row-icon bad">${icon('exclamation.triangle.fill')}</span>
      <div class="row-main"><span class="row-title"><b>הסנכרון עם GitHub לא עובד</b></span>
        <span class="row-sub">${esc(e.message)}. העבודה שמורה במכשיר הזה.</span></div>
      <button class="btn btn-tinted btn-small" data-go="files">פרטים</button>
    </div></div>`;
  }
  if (syncAuto() || !syncFlags().news) return '';
  return `
    <div class="list"><div class="notice">
      <span class="row-icon">${icon('square.and.arrow.down')}</span>
      <div class="row-main"><span class="row-title"><b>יש עדכון ב-GitHub</b></span>
        <span class="row-sub">${esc(`מ-${isolate(syncCfg.remoteDevice || 'מכשיר אחר')}${syncCfg.remoteAt ? ` · ${whenLabel(syncCfg.remoteAt)}` : ''}. עוד לא יובא למכשיר הזה.`)}</span></div>
      <button class="btn btn-tinted btn-small" data-act="sync-pull">ייבוא</button>
    </div></div>`;
}

/** On GitHub Pages: the repository the site itself is served from. */
function siteRepo() {
  const m = location.hostname.match(/^([\w-]+)\.github\.io$/i);
  if (!m) return null;
  const first = location.pathname.split('/').filter(Boolean)[0];
  return { owner: m[1], repo: first && !/\.html?$/i.test(first) ? first : `${m[1]}.github.io` };
}
const sameRepo = (a, b) => !!(a && b) && a.owner.toLowerCase() === b.owner.toLowerCase() && a.repo.toLowerCase() === b.repo.toLowerCase();

/** Check the repository, keep the settings on this device, then import and save once. */
async function connectSync(cfg) {
  const say = (text, bad = false) => { const el = $('#gh-result'); if (el) { el.textContent = text; el.classList.toggle('bad', bad); } };
  if (sameRepo(cfg, siteRepo())) {
    say('זה הריפו של האתר עצמו, והוא נשאר ציבורי בשביל GitHub Pages. לנתונים צריך ריפו נוסף ונפרד, פרטי.', true);
    return;
  }
  say('בודק את החיבור ל-GitHub…');
  try {
    await checkRepo(cfg);
  } catch (e) {
    say(e.message, true);
    return;
  }
  syncCfg = { owner: cfg.owner, repo: cfg.repo, path: cfg.path || DEFAULT_PATH, token: cfg.token, auto: false };
  saveSyncCfg();
  syncState.error = null;
  closeDialog();
  render();
  await runSync('both', { manual: true });
}

function syncSetup() {
  const site = siteRepo();
  const example = `${site ? site.owner : 'שם-משתמש'}/shibutz-data`;
  const newRepo = `https://github.com/new?name=shibutz-data&visibility=private&description=${encodeURIComponent('נתוני שיבוץ המדריכים')}`;
  openSheet({
    title: 'סנכרון דרך GitHub',
    sub: 'האתר נשאר בריפו הציבורי שלו, ו-GitHub Pages ממשיך לעבוד. הנתונים נשמרים בריפו שני, נפרד ופרטי.',
    done: 'חיבור',
    body: `
      <section class="list-section">
        <h3 class="list-header">פעם אחת, באתר GitHub</h3>
        <div class="list">
          <a class="row action" href="${newRepo}" target="_blank" rel="noopener">${icon('plus.circle.fill')}<span class="row-main"><span class="row-title">1. ריפו נוסף, רק לנתונים</span><span class="row-sub">לא הריפו של האתר. שם לדוגמה: shibutz-data · לבחור Private (בחשבון החינמי ריפו פרטי הוא בחינם)</span></span>${icon('arrow.up.forward', 'chevron')}</a>
          <a class="row action" href="https://github.com/settings/personal-access-tokens/new" target="_blank" rel="noopener">${icon('key')}<span class="row-main"><span class="row-title">2. טוקן לריפו הזה</span><span class="row-sub">Repository access: Only select repositories ← ריפו הנתונים. Permissions ← Contents: Read and write.</span></span>${icon('arrow.up.forward', 'chevron')}</a>
        </div>
      </section>
      <section class="list-section">
        <h3 class="list-header">3. במכשיר הזה</h3>
        <div class="list">
          <div class="row"><input class="field-inline" id="gh-repo" placeholder="ריפו הנתונים, למשל ${esc(example)}" autocomplete="off" autocapitalize="off" spellcheck="false" aria-label="ריפו הנתונים"></div>
          <div class="row"><input class="field-inline" id="gh-token" type="password" placeholder="הטוקן (github_pat_…)" autocomplete="off" spellcheck="false" aria-label="הטוקן"></div>
        </div>
        <p class="list-footer" id="gh-result">בריפו נשמר קובץ אחד, ${DEFAULT_PATH}, ו-GitHub שומר כל גרסה שלו. בקובץ יש את פרטי המדריכים, ולכן הריפו חייב להיות פרטי. הטוקן נשמר רק במכשיר הזה. אפשר גם להדביק כאן קישור חיבור ממכשיר אחר.</p>
      </section>`,
  }, {
    save() {
      const raw = $('#gh-repo').value.trim();
      const fromLink = readConnectLink(raw);
      if (fromLink) { connectSync(fromLink); return; }
      const repo = parseRepo(raw);
      const token = $('#gh-token').value.trim();
      if (!repo || !token) {
        const el = $('#gh-result');
        el.textContent = 'צריך את שם ריפו הנתונים (שם-משתמש/שם-הריפו) ואת הטוקן.';
        el.classList.add('bad');
        return;
      }
      connectSync({ ...repo, token });
    },
  });
}

/** Opened from a connect link (QR or message): ask before connecting. */
function confirmConnect(cfg) {
  openSheet({
    title: 'חיבור לסנכרון',
    sub: `המכשיר הזה יתחבר לריפו ${esc(`${cfg.owner}/${cfg.repo}`)} ב-GitHub.`,
    done: 'חיבור',
    body: `<p class="list-footer" id="gh-result" style="padding-block-start:0">מה שנשמר שם ייובא לכאן, ומה שכבר יש במכשיר הזה יתמזג איתו. שום שיבוץ לא יימחק.</p>`,
  }, { save() { connectSync(cfg); } });
}

async function syncShare() {
  const link = connectLink(syncCfg, location.href.split('#')[0]);
  openSheet({
    title: 'חיבור מכשיר נוסף',
    sub: 'סורקים את הקוד במצלמה של הטלפון, או שולחים לעצמכם את הקישור ופותחים אותו במכשיר השני.',
    body: `
      <div class="qr-box"><div id="qr" class="qr" role="img" aria-label="קוד QR לחיבור מכשיר נוסף"></div></div>
      <div class="list"><button class="row action" data-act="copy-connect">${icon('link')}<span class="row-main">העתקת קישור החיבור</span></button></div>
      <p class="list-footer">בקישור יש את הטוקן: מי שמחזיק בו יכול לקרוא ולשנות את קובץ הנתונים, אז שולחים אותו רק לעצמכם. אם האתר מותקן במסך הבית של האייפון, פותחים אותו משם ומדביקים את הקישור ב"סנכרון דרך GitHub".</p>`,
  }, { link });
  try {
    const qrcode = await loadScript(QR_URL, 'qrcode');
    const qr = qrcode(0, 'L');
    qr.addData(link);
    qr.make();
    const box = $('#qr');
    if (box) box.innerHTML = qr.createSvgTag({ cellSize: 4, margin: 2, scalable: true });
  } catch {
    const box = $('#qr');
    if (box) box.textContent = 'קוד ה-QR לא נטען. אפשר להעתיק את הקישור.';
  }
}

/** Delete the data file from GitHub. This device keeps its data; automatic sync is turned off so it does not write it back. */
async function syncDeleteRemote() {
  if (!syncOn()) return;
  if (syncState.busy) { toast('רגע - פעולה אחרת מול GitHub עדיין רצה'); return; }
  const cfg = syncCfg;
  const wasAuto = !!cfg.auto;
  clearTimeout(syncState.timer);
  Object.assign(syncState, { queued: null, pending: null, busy: 'delete' });
  renderSyncBits();
  try {
    let deleted = false;
    for (let attempt = 0; attempt < 3; attempt++) {
      const remote = await readFile(cfg);
      if (!remote) break;
      try {
        deleted = await deleteFile(cfg, remote.sha, `מחיקת הנתונים מ-${device.name}`);
        break;
      } catch (e) {
        if (e.code !== 'conflict' || attempt === 2) throw e; // someone saved in between: read again and delete that
      }
    }
    Object.assign(cfg, { auto: false, sha: null, etag: null, remoteHash: null, remoteDevice: null, remoteAt: null, importedSha: null, cleanHash: null });
    syncState.error = null;
    toast(deleted
      ? `הנתונים נמחקו מ-GitHub. במכשיר הזה הם נשארו${wasAuto ? ', והסנכרון האוטומטי כובה' : ''}.`
      : 'ב-GitHub כבר לא היו נתונים.');
  } catch (e) {
    syncState.error = e;
    toast('המחיקה נכשלה: ' + e.message);
  } finally {
    if (syncCfg === cfg) saveSyncCfg();
    syncState.busy = null;
    render();
  }
}

function syncOff() {
  syncCfg = null;
  saveSyncCfg();
  syncState.error = null;
  render();
  toast('הסנכרון נותק במכשיר הזה. הנתונים נשארו כאן וגם ב-GitHub.');
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
  $('#nav-leading').innerHTML = syncButtonHtml();
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
    if (t.closest('[data-act="ev-menu"], [data-act="sync-menu"]')) return; // the same button closes its menu
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
    if (act === 'copy-connect') { copyText(dialogHandlers.link); return; }
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
    case 'send': sendToDevice(); break;
    case 'sync-setup': syncSetup(); break;
    case 'sync-menu': syncMenu(actEl); break;
    case 'sync-push': runSync('push', { manual: true }); break;
    case 'sync-pull': runSync('pull', { manual: true }); break;
    case 'sync-both': runSync('both', { manual: true }); break;
    case 'sync-delete':
      openAlert({
        title: 'למחוק את הנתונים מ-GitHub?',
        message: `קובץ הנתונים יימחק מהריפו ${isolate(`${syncCfg.owner}/${syncCfg.repo}`)}. במכשיר הזה הנתונים נשארים${syncCfg.auto ? ', והסנכרון האוטומטי ייכבה' : ''}. מכשיר אחר שעדיין מחובר יעלה אותם שוב אם ילחצו בו "שמירה" או אם מופעל בו סנכרון אוטומטי. GitHub שומר גרסאות קודמות בהיסטוריה של הריפו; למחיקה מוחלטת מוחקים את הריפו עצמו בהגדרות שלו.`,
        actions: [
          { label: 'ביטול', cls: 'btn-gray' },
          { label: 'מחיקה', cls: 'btn-prominent btn-destructive', run: syncDeleteRemote },
        ],
      });
      break;
    case 'sync-share': syncShare(); break;
    case 'sync-off': syncOff(); break;
    case 'reset':
      openAlert({
        title: 'למחוק את כל הנתונים?',
        message: syncOn()
          ? 'המדריכים, האירועים, השיבוצים וההערות יימחקו מהמכשיר הזה, והסנכרון בו ינותק. הקובץ ב-GitHub נשאר כמו שהוא.'
          : 'המדריכים, האירועים, השיבוצים וההערות יימחקו מהמכשיר הזה. אי אפשר לבטל את זה - כדאי לשלוח קודם עותק למכשיר אחר.',
        actions: [
          { label: 'ביטול', cls: 'btn-gray' },
          { label: 'מחיקה', cls: 'btn-prominent btn-destructive', run: () => {
            if (syncCfg) { syncCfg = null; saveSyncCfg(); syncState.error = null; }
            Object.assign(device, { sentHash: null, received: null });
            saveDevice();
            state = emptyState();
            snapshot = clone(state);
            autoUndo = null;
            notices.roster = notices.schedule = null;
            toast('כל הנתונים נמחקו מהמכשיר הזה');
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
  if (t.id === 'file-receive' && t.files[0]) receiveFile(t.files[0]);
  if (t.id === 'sync-auto' && syncCfg) {
    syncCfg.auto = t.checked;
    saveSyncCfg();
    toast(t.checked ? 'סנכרון אוטומטי פועל: כל שינוי נשמר ב-GitHub תוך שניות, ועדכונים מיובאים לבד.' : 'סנכרון אוטומטי כבוי: שומרים ומייבאים בלחיצה.');
    if (t.checked) scheduleSync('both', 300);
    return;
  }
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
      persist();
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
  if (/\.(json|txt)$/i.test(file.name)) receiveFile(file);
  else if (zone.dataset.drop === 'roster') importRoster(file);
  else importSchedule(file);
});

$('#dialog').addEventListener('close', () => { dialogHandlers = {}; resumePending(); });
$('#alert').addEventListener('close', () => resumePending());
document.addEventListener('focusout', () => resumePending(1500));
window.addEventListener('online', () => scheduleSync(syncAuto() ? 'both' : 'check', 500));
document.addEventListener('toggle', e => { if (e.target.id === 'filters-box') filtersOpen = e.target.open; }, true);
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') { flush(); if (syncAuto() && syncFlags().unsaved) runSync('both'); } else scheduleSync(syncAuto() ? 'both' : 'check', 300);
});
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
  snapshot = clone(state);
  syncState.localHash = demo ? null : sharedHash(state);
  personIndex = null;
  render();
  if (hasData(state) && !demo) askToPersist();
  // A connect link (QR code or message) opens the sync confirmation once, then leaves the address bar.
  const link = readConnectLink(location.hash);
  if (link) {
    history.replaceState(null, '', location.pathname + location.search);
    if (store.kind === 'local') confirmConnect(link);
  } else if (syncOn()) runSync(syncAuto() ? 'both' : 'check');
  // Without automatic sync this only looks whether another device saved something, to show it.
  setInterval(() => { if (document.visibilityState === 'visible') runSync(syncAuto() ? 'both' : 'check'); }, 90000);
}

start();
