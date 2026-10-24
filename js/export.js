// The output document, in the layout of the "משימות" file: what is staffed, what
// is still open (and who has not been assigned yet), courses, and past sessions.
// One block model feeds the Word file, the on-screen preview and the plain text.

import { DAY_NAMES, weekday } from './roster.js';
import { visibleSlots, isRelevant, allAssignments } from './scheduling.js';
import { normalizeName } from './names.js';

export const MONTH_NAMES = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];

export function shortDate(iso) {
  const [, m, d] = iso.split('-');
  return `${d}/${m}`;
}

export function dayLabel(iso) {
  return `יום ${DAY_NAMES[weekday(iso)]}- ${shortDate(iso)}`;
}

export function monthLabel(key, withYear) {
  const [y, m] = key.split('-');
  return MONTH_NAMES[+m - 1] + (withYear ? ` ${y}` : '');
}

const text = (t, extra = {}) => ({ text: t, ...extra });
const missing = () => text('?', { missing: true });

function roleWord(slot, count) {
  if (slot.kind === 'lecture') return 'מרצה';
  return count > 1 ? 'מדריכים' : 'מדריך';
}

function staffRuns(positions) {
  const runs = [];
  positions.forEach((p, i) => {
    if (i) runs.push(text(' + '));
    runs.push(p.name ? text(p.name) : missing());
  });
  return runs;
}

const endDot = s => (/[.!?:]$/.test(s) ? s : s + '.');

/** The lines of one slot: "10:00-11:30 - label. מדריכים: X + Y." (+ one line per labelled group). */
export function slotLines(slot, options = {}) {
  const lines = [];
  const labelled = slot.groups.some(g => g.label);
  const first = [text(`${slot.start}-${slot.end} - ${endDot(slot.label)}`)];
  if (!labelled) {
    const count = slot.groups.reduce((n, g) => n + g.positions.length, 0);
    first.push(text(` ${roleWord(slot, count)}: `));
    slot.groups.forEach((g, i) => {
      if (i) first.push(text(', '));
      first.push(...staffRuns(g.positions));
    });
    first.push(text('.'));
  }
  const groupRuns = g => {
    const activity = g.activity && g.activity !== slot.label ? ` - ${g.activity}` : '';
    return [text(`${g.label || 'קבוצה'}${activity}- ${roleWord(slot, g.positions.length)}: `), ...staffRuns(g.positions), text('.')];
  };
  // A single labelled group stays on the slot's line.
  if (labelled && slot.groups.length === 1) first.push(text(' '), ...groupRuns(slot.groups[0]));
  if (options.location && slot.location) first.push(text(` מיקום: ${endDot(slot.location)}`));
  lines.push(first);
  if (labelled && slot.groups.length > 1) for (const g of slot.groups) lines.push(groupRuns(g));
  return lines;
}

function eventBlocks(ev, settings, options) {
  const blocks = [];
  if (ev.title) blocks.push({ type: 'p', runs: [text(ev.title)] });
  for (const slot of visibleSlots(ev, settings)) {
    for (const runs of slotLines(slot, options)) blocks.push({ type: 'p', runs });
  }
  if (ev.note) blocks.push({ type: 'bullet', runs: [text(ev.note)] });
  return blocks;
}

const isOpen = (ev, settings) => visibleSlots(ev, settings).some(s => s.groups.some(g => g.positions.some(p => !p.name)));

function byMonth(events, blocksFor) {
  const blocks = [];
  const months = [...new Set(events.map(e => e.date.slice(0, 7)))];
  const years = new Set(months.map(m => m.slice(0, 4)));
  for (const month of months) {
    blocks.push({ type: 'h2', runs: [text(monthLabel(month, years.size > 1))] });
    let lastDate = null;
    for (const ev of events.filter(e => e.date.startsWith(month))) {
      if (ev.date !== lastDate) {
        if (lastDate) blocks.push({ type: 'blank' });
        blocks.push({ type: 'p', runs: [text(`${dayLabel(ev.date)}-`, { bold: true })] });
        lastDate = ev.date;
      } else blocks.push({ type: 'blank' });
      blocks.push(...blocksFor(ev));
    }
  }
  return blocks;
}

/** "סביון שכבה ח (100 תל')- מפגש 1." -> "סביון שכבה ח" */
export function courseName(title) {
  return title.replace(/\s*[-–]?\s*מפגש.*$/, '').replace(/\s*\(.*$/, '').replace(/[\s\-–.]+$/, '').trim() || title;
}

/**
 * @param state    {events, instructors, notes, settings}
 * @param options  {month: 'YYYY-MM'|'all', today: 'YYYY-MM-DD', sections: {assigned, open, pool, courses, done}, location: bool}
 * @returns blocks [{type: 'h1'|'h2'|'p'|'bullet'|'blank', runs: [{text, bold?, missing?}]}]
 */
export function buildDocument(state, options) {
  const { settings } = state;
  const sections = options.sections || {};
  const inMonth = ev => !options.month || options.month === 'all' || ev.date.startsWith(options.month);
  const events = state.events
    .filter(ev => !ev.ignored && isRelevant(ev) && settings.kinds[ev.kind] !== false && visibleSlots(ev, settings).length && inMonth(ev))
    .sort((a, b) => a.date.localeCompare(b.date));
  const today = options.today;
  const upcoming = events.filter(ev => !today || ev.date >= today);
  const past = events.filter(ev => today && ev.date < today);
  const blocks = [];
  const bullets = t => String(t || '').split('\n').map(s => s.replace(/^[\s\-•*]+/, '').trim()).filter(Boolean)
    .map(s => ({ type: 'bullet', runs: [text(s)] }));

  if (state.notes?.gaps?.trim()) {
    blocks.push({ type: 'p', runs: [text('פערים לטיפול:', { bold: true })] }, ...bullets(state.notes.gaps));
  }

  if (sections.assigned !== false) {
    const done = upcoming.filter(ev => !isOpen(ev, settings));
    blocks.push({ type: 'h1', runs: [text('שובץ')] });
    blocks.push(...(done.length ? byMonth(done, ev => eventBlocks(ev, settings, options)) : [{ type: 'p', runs: [text('אין עדיין אירועים משובצים במלואם.')] }]));
  }

  if (sections.open !== false) {
    const open = upcoming.filter(ev => isOpen(ev, settings));
    blocks.push({ type: 'h1', runs: [text('לא שובץ')] });
    blocks.push(...(open.length ? byMonth(open, ev => eventBlocks(ev, settings, options)) : [{ type: 'p', runs: [text('כל האירועים שובצו.')] }]));

    if (sections.pool !== false) {
      const assigned = new Set(allAssignments(upcoming).map(a => a.key));
      const free = role => state.instructors
        .filter(i => i.active !== false && i.roles?.[role] && !assigned.has(normalizeName(i.name)))
        .map(i => i.name);
      const pools = [['הרצאות', free('lecture')], ['מופע קסמים', free('magic')], ['הדרכות ימי עיון וקורסים', free('lab')]];
      if (pools.some(([, names]) => names.length)) {
        blocks.push({ type: 'blank' }, { type: 'p', runs: [text('מדריכים שעוד לא שובצו:', { bold: true })] });
        for (const [title, names] of pools) {
          if (!names.length) continue;
          blocks.push({ type: 'bullet', runs: [text(`${title}: `, { bold: true }), text(names.join(', '))] });
        }
      }
    }
  }

  if (sections.courses !== false) {
    const courses = new Map();
    for (const ev of events.filter(e => e.kind === 'course')) {
      const name = courseName(ev.title);
      if (!courses.has(name)) courses.set(name, []);
      courses.get(name).push(ev);
    }
    if (courses.size) {
      blocks.push({ type: 'h1', runs: [text('קורסים')] });
      for (const [name, list] of courses) {
        blocks.push({ type: 'p', runs: [text(`${name}:`, { bold: true })] });
        for (const ev of list) {
          const runs = [text(`${shortDate(ev.date)}- `)];
          visibleSlots(ev, settings).forEach((s, si) => {
            if (si) runs.push(text(' | '));
            if (visibleSlots(ev, settings).length > 1) runs.push(text(`${s.label}: `));
            s.groups.forEach((g, gi) => { if (gi) runs.push(text(', ')); runs.push(...staffRuns(g.positions)); });
          });
          blocks.push({ type: 'bullet', runs });
        }
      }
    }
  }

  if (sections.done !== false && past.length) {
    blocks.push({ type: 'h1', runs: [text('הדרכות שבוצעו')] });
    let first = true;
    for (const ev of past) {
      if (!first) blocks.push({ type: 'blank' });
      first = false;
      blocks.push({ type: 'p', runs: [text(`${shortDate(ev.date)}-`, { bold: true })] });
      for (const s of visibleSlots(ev, settings)) {
        const runs = [text(`${s.label}- `)];
        s.groups.forEach((g, gi) => {
          if (gi) runs.push(text(', '));
          g.positions.forEach((p, pi) => { if (pi) runs.push(text(', ')); runs.push(p.name ? text(p.name) : missing()); });
        });
        blocks.push({ type: 'p', runs });
      }
    }
  }

  if (state.notes?.tasks?.trim()) {
    blocks.push({ type: 'blank' }, { type: 'p', runs: [text('משימות', { bold: true })] }, ...bullets(state.notes.tasks));
  }
  return blocks;
}

export function toPlainText(blocks) {
  return blocks.map(b => {
    const t = b.runs ? b.runs.map(r => r.text).join('') : '';
    if (b.type === 'h1') return `\n${t}\n${'='.repeat(Math.min(t.length, 20))}`;
    if (b.type === 'h2') return `\n${t}`;
    if (b.type === 'bullet') return `• ${t}`;
    if (b.type === 'blank') return '';
    return t;
  }).join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

export function toHtml(blocks) {
  const runs = rs => rs.map(r => {
    let h = esc(r.text);
    if (r.missing) h = `<mark>${h}</mark>`;
    if (r.bold) h = `<b>${h}</b>`;
    return h;
  }).join('');
  let html = '';
  let inList = false;
  for (const b of blocks) {
    if (b.type === 'bullet' && !inList) { html += '<ul>'; inList = true; }
    if (b.type !== 'bullet' && inList) { html += '</ul>'; inList = false; }
    if (b.type === 'h1') html += `<h1>${runs(b.runs)}</h1>`;
    else if (b.type === 'h2') html += `<h2>${runs(b.runs)}</h2>`;
    else if (b.type === 'bullet') html += `<li>${runs(b.runs)}</li>`;
    else if (b.type === 'blank') html += '<p class="blank">&nbsp;</p>';
    else html += `<p>${runs(b.runs)}</p>`;
  }
  if (inList) html += '</ul>';
  return html;
}

/** Word document (right-to-left) built with the docx library. */
export async function toDocxBlob(docx, blocks, title) {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel } = docx;
  const runs = rs => rs.map(r => new TextRun({
    text: r.text, rightToLeft: true, bold: r.bold, highlight: r.missing ? 'yellow' : undefined,
  }));
  const children = blocks.map(b => {
    if (b.type === 'h1') return new Paragraph({ heading: HeadingLevel.HEADING_1, bidirectional: true, children: runs(b.runs), spacing: { before: 240, after: 120 } });
    if (b.type === 'h2') return new Paragraph({ heading: HeadingLevel.HEADING_2, bidirectional: true, children: runs(b.runs), spacing: { before: 200, after: 80 } });
    if (b.type === 'bullet') return new Paragraph({ bullet: { level: 0 }, bidirectional: true, children: runs(b.runs) });
    if (b.type === 'blank') return new Paragraph({ bidirectional: true, children: [] });
    return new Paragraph({ bidirectional: true, children: runs(b.runs) });
  });
  const doc = new Document({
    title,
    creator: 'שיבוץ מדריכים',
    styles: { default: { document: { run: { font: 'Arial', size: 22, rightToLeft: true } } } },
    sections: [{ properties: {}, children }],
  });
  return Packer.toBlob(doc);
}
