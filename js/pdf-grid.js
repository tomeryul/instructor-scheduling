// Reads a calendar-style schedule PDF (an Excel sheet saved as PDF: weeks of 7 day
// columns, with a row of dates at the top of each week) and returns the text of
// every cell together with its date and background colour.
//
// Office draws every text run inside a clip rectangle equal to its cell, so the
// clip tells us which cell a piece of text belongs to - even when the text
// overflows into the cell below it.

const DATE_RE = /^(\d{1,2})[./](\d{1,2})[./](\d{2}|\d{4})$/;
const HEBREW_RE = /[֐-׿]/;
const STRONG_LTR_RE = /[0-9A-Za-z]/;

// pdf.js path operators inside constructPath
const PATH_MOVE = 13, PATH_LINE = 14, PATH_CURVE = 15, PATH_CURVE2 = 16, PATH_CURVE3 = 17, PATH_RECT = 19;

const mul = (m, n) => [
  m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5],
];
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

function boxOf(m, x0, y0, x1, y1) {
  const pts = [apply(m, x0, y0), apply(m, x1, y0), apply(m, x0, y1), apply(m, x1, y1)];
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
}

function toHex(c) {
  if (typeof c === 'string') return c.toLowerCase();
  return '#' + [c[0], c[1], c[2]].map(v => Math.round(v).toString(16).padStart(2, '0')).join('');
}

function pathBox(ops, coords) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity, k = 0;
  const add = (x, y) => { x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y); };
  for (const op of ops) {
    if (op === PATH_RECT) { const [x, y, w, h] = coords.slice(k, k + 4); add(x, y); add(x + w, y + h); k += 4; }
    else if (op === PATH_MOVE || op === PATH_LINE) { add(coords[k], coords[k + 1]); k += 2; }
    else if (op === PATH_CURVE) { add(coords[k + 4], coords[k + 5]); k += 6; }
    else if (op === PATH_CURVE2 || op === PATH_CURVE3) { add(coords[k + 2], coords[k + 3]); k += 4; }
  }
  return isFinite(x0) ? { x0, y0, x1, y1 } : null;
}

/** Walk the drawing operators: filled boxes (cell colours, borders) and the clip of every text run. */
async function readGraphics(page, OPS) {
  const ol = await page.getOperatorList();
  const fills = [];
  const runs = [];
  let ctm = [1, 0, 0, 1, 0, 0];
  let clip = null;
  let fill = '#000000';
  let path = null;
  const stack = [];
  for (let i = 0; i < ol.fnArray.length; i++) {
    const fn = ol.fnArray[i];
    const args = ol.argsArray[i];
    switch (fn) {
      case OPS.save: stack.push({ ctm, clip, fill }); break;
      case OPS.restore: if (stack.length) ({ ctm, clip, fill } = stack.pop()); break;
      case OPS.transform: ctm = mul(args, ctm); break;
      case OPS.setFillRGBColor: fill = toHex(args); break;
      case OPS.setFillGray: { const v = args[0] * 255; fill = toHex([v, v, v]); break; }
      case OPS.constructPath: {
        const b = pathBox(args[0], args[1]);
        path = b && boxOf(ctm, b.x0, b.y0, b.x1, b.y1);
        break;
      }
      case OPS.clip: case OPS.eoClip:
        if (path) clip = path;
        break;
      case OPS.fill: case OPS.eoFill: case OPS.fillStroke: case OPS.eoFillStroke:
        if (path) fills.push({ ...path, color: fill });
        path = null;
        break;
      case OPS.endPath: path = null; break;
      case OPS.setTextMatrix: {
        const [x, y] = apply(ctm, args[4], args[5]);
        runs.push({ x, y, clip });
        break;
      }
      default: break;
    }
  }
  return { fills, runs };
}

function charDir(str) {
  if (HEBREW_RE.test(str)) return 'rtl';
  if (STRONG_LTR_RE.test(str)) return 'ltr';
  return 'n';
}

/** Turn the items of one visual line into reading-order text (base direction right-to-left). */
export function lineToText(items, fontSize) {
  const vis = [...items].sort((a, b) => a.x0 - b.x0);
  const dirs = vis.map(it => charDir(it.str));
  // A neutral item between two LTR items (the '-' in 050-1234567) belongs to the LTR run.
  for (let i = 0; i < vis.length; i++) {
    if (dirs[i] !== 'n') continue;
    let l = i - 1; while (l >= 0 && dirs[l] === 'n') l--;
    let r = i + 1; while (r < vis.length && dirs[r] === 'n') r++;
    dirs[i] = l >= 0 && r < vis.length && dirs[l] === 'ltr' && dirs[r] === 'ltr' ? 'ltr' : 'rtl';
  }
  const runs = [];
  vis.forEach((it, i) => {
    const last = runs[runs.length - 1];
    if (last && last.dir === dirs[i]) last.items.push(it);
    else runs.push({ dir: dirs[i], items: [it] });
  });
  const gap = (left, right) => (right.x0 - left.x1 > fontSize * 0.12 ? ' ' : '');
  let out = '';
  for (let r = runs.length - 1; r >= 0; r--) {
    const run = runs[r];
    if (r < runs.length - 1) out += gap(run.items[run.items.length - 1], runs[r + 1].items[0]);
    const seq = run.dir === 'ltr' ? run.items : [...run.items].reverse();
    seq.forEach((it, j) => {
      if (j > 0) out += run.dir === 'ltr' ? gap(seq[j - 1], it) : gap(it, seq[j - 1]);
      out += it.str;
    });
  }
  return out.replace(/\s+/g, ' ').trim();
}

function groupLines(items) {
  const lines = [];
  for (const it of [...items].sort((a, b) => b.y - a.y)) {
    const line = lines[lines.length - 1];
    if (line && Math.abs(line.y - it.y) < Math.max(line.size, it.size) * 0.35) line.items.push(it);
    else lines.push({ y: it.y, size: it.size, items: [it] });
  }
  return lines;
}

/**
 * Excel wraps long text, so a line that reaches the far (left) edge of the cell
 * continues on the next one; a line that stops short ended with a real line break.
 */
function joinLines(lines, cellLeft) {
  const out = [];
  let wrapped = false;
  lines.forEach((line, i) => {
    const text = lineToText(line.items, line.size);
    if (!text) return;
    if (wrapped && out.length) {
      const prev = out[out.length - 1];
      out[out.length - 1] = prev + (/[-–/]$/.test(prev) && /^\d/.test(text) ? '' : ' ') + text;
    } else out.push(text);
    const next = lines[i + 1];
    if (!next) return;
    const left = Math.min(...line.items.map(it => it.x0));
    const nextText = lineToText(next.items, next.size);
    const nextWidth = Math.max(...next.items.map(it => it.x1)) - Math.min(...next.items.map(it => it.x0));
    const firstWord = nextText.split(' ')[0] || '';
    const wordWidth = nextText.length ? (nextWidth * firstWord.length) / nextText.length : 0;
    wrapped = left - cellLeft < wordWidth + line.size * 0.8;
  });
  return out;
}

export function parseDateText(str) {
  const m = String(str).trim().match(DATE_RE);
  if (!m) return null;
  let year = +m[3];
  if (year < 100) year += 2000;
  return `${year}-${String(+m[2]).padStart(2, '0')}-${String(+m[1]).padStart(2, '0')}`;
}

const boxKey = b => [b.x0, b.y0, b.x1, b.y1].map(v => v.toFixed(1)).join(',');

/**
 * @param pdfjsLib  the pdf.js module (getDocument, OPS)
 * @param data      Uint8Array with the PDF bytes
 * @returns {Promise<{cells: {date: string, lines: string[], color: string|null, page: number}[], warnings: string[]}>}
 */
export async function extractScheduleCells(pdfjsLib, data) {
  const doc = await pdfjsLib.getDocument({ data, isEvalSupported: false }).promise;
  const cells = [];
  const warnings = [];
  let lastWeek = null; // a week can continue on the next page

  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    const items = tc.items
      .filter(it => it.str && it.str.trim())
      .map(it => {
        const x0 = it.transform[4];
        const size = Math.hypot(it.transform[2], it.transform[3]) || it.height || 1;
        return { str: it.str, x0, x1: x0 + it.width, y: it.transform[5], size };
      });
    if (!items.length) continue;
    const { fills, runs } = await readGraphics(page, pdfjsLib.OPS);

    // Borders are thin filled rectangles.
    const fallbackSize = Math.min(...items.map(i => i.size));
    const thin = fallbackSize * 0.5;
    const vXs = fills.filter(f => f.x1 - f.x0 < thin && f.y1 - f.y0 > thin).map(f => (f.x0 + f.x1) / 2);
    const hRules = fills.filter(f => f.y1 - f.y0 < thin && f.x1 - f.x0 > thin);

    // Week headers: three or more dates on one baseline. Columns run right to left.
    const dateRows = [];
    for (const it of items.filter(i => DATE_RE.test(i.str.trim())).sort((a, b) => b.y - a.y)) {
      const row = dateRows[dateRows.length - 1];
      if (row && Math.abs(row.y - it.y) < it.size * 0.5) row.items.push(it);
      else dateRows.push({ y: it.y, items: [it] });
    }
    const weeks = dateRows.filter(r => r.items.length >= 3).map(r => {
      const cols = r.items.map(it => ({ date: parseDateText(it.str), cx: (it.x0 + it.x1) / 2 })).sort((a, b) => b.cx - a.cx);
      cols.forEach((c, i) => {
        const left = vXs.filter(x => x < c.cx), right = vXs.filter(x => x > c.cx);
        c.x0 = left.length ? Math.max(...left) : i < cols.length - 1 ? (c.cx + cols[i + 1].cx) / 2 : -Infinity;
        c.x1 = right.length ? Math.min(...right) : i > 0 ? (c.cx + cols[i - 1].cx) / 2 : Infinity;
      });
      return { y: r.y, cols };
    });
    // The whole date row (Hebrew dates included) is a header, not a cell.
    const headerItems = new Set(items.filter(it => weeks.some(w => Math.abs(w.y - it.y) < it.size * 0.5)));
    const colWidth = weeks.length ? Math.max(...weeks[0].cols.map(c => c.x1 - c.x0).filter(isFinite)) : Infinity;

    // Which cell does each text item belong to? The clip of the run it was drawn in
    // (on the same baseline, the closest run that starts at or before the item), as
    // long as that clip is a cell and not the whole table.
    for (const it of items) {
      const eps = it.size * 0.05;
      let best = null;
      for (const r of runs) {
        if (!r.clip || Math.abs(r.y - it.y) > eps || r.x > it.x0 + eps) continue;
        if (!best || r.x > best.x) best = r;
      }
      it.clip = best && best.clip.x1 - best.clip.x0 < colWidth * 2.5 ? best.clip : null;
    }

    // Unclipped text: the cell is bounded by the column and the row rules around it.
    const rowOf = (col, y) => {
      const ys = hRules.filter(r => r.x0 <= col.x0 + thin && r.x1 >= col.x1 - thin).map(r => (r.y0 + r.y1) / 2);
      const above = ys.filter(v => v > y), below = ys.filter(v => v < y);
      return { y1: above.length ? Math.min(...above) : y + fallbackSize, y0: below.length ? Math.max(...below) : y - fallbackSize };
    };

    const groups = new Map();
    for (const it of items) {
      if (headerItems.has(it)) continue;
      // The week whose date row is the nearest one above the cell; above the first
      // date row of a page, the week that was running at the end of the previous page.
      const top = it.clip ? it.clip.y1 - 0.01 : it.y;
      const owner = weeks.filter(w => w.y > top).pop() || lastWeek;
      if (!owner) continue;
      const cx = it.clip ? (it.clip.x0 + it.clip.x1) / 2 : (it.x0 + it.x1) / 2;
      const col = owner.cols.find(c => cx >= c.x0 && cx <= c.x1);
      if (!col || !col.date) continue;
      const box = it.clip || { x0: col.x0, x1: col.x1, ...rowOf(col, it.y) };
      const key = boxKey(box);
      if (!groups.has(key)) groups.set(key, { date: col.date, box, items: [] });
      groups.get(key).items.push(it);
    }

    for (const g of groups.values()) {
      const box = g.box;
      const left = isFinite(box.x0) ? box.x0 : Math.min(...g.items.map(i => i.x0));
      const lines = joinLines(groupLines(g.items), left);
      if (!lines.length) continue;
      const mx = (box.x0 + box.x1) / 2, my = (box.y0 + box.y1) / 2;
      const bg = fills
        .filter(f => f.x1 - f.x0 > fallbackSize && f.y1 - f.y0 > fallbackSize && f.x0 <= mx && f.x1 >= mx && f.y0 <= my && f.y1 >= my)
        .sort((a, b) => (a.x1 - a.x0) * (a.y1 - a.y0) - (b.x1 - b.x0) * (b.y1 - b.y0))[0];
      cells.push({ date: g.date, lines, color: bg && bg.color !== '#ffffff' ? bg.color : null, page: p, top: box.y1, x: mx });
    }
    if (weeks.length) lastWeek = weeks[weeks.length - 1];
    else if (!lastWeek) warnings.push(`עמוד ${p}: לא נמצאה שורת תאריכים`);
  }

  // Reading order: by date, then top to bottom.
  cells.sort((a, b) => a.date.localeCompare(b.date) || a.page - b.page || b.top - a.top);
  return { cells: cells.map(({ top, x, ...c }) => c), warnings };
}
