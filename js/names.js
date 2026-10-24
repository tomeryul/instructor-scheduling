// Hebrew name normalisation and fuzzy matching against the instructor roster.
// The schedule spells names a little differently from the roster
// ("היבה" / "היבא", "בוצ'קו" / "בוצקו", "יעל בלייברג" / "יעל שא-נס בלייברג").

const FINALS = { 'ך': 'כ', 'ם': 'מ', 'ן': 'נ', 'ף': 'פ', 'ץ': 'צ' };

export function normalizeName(s) {
  return String(s || '')
    .replace(/[֑-ׇ]/g, '')
    .replace(/["'`׳״?]/g, '')
    .replace(/[-–—_.,:;()]/g, ' ')
    .replace(/[ךםןףץ]/g, c => FINALS[c])
    .split(/\s+/)
    .filter(Boolean)
    .map(w => (w.length > 2 ? w.replace(/א$/, 'ה') : w))
    .join(' ');
}

/** Clean a name for display: trims spaces and stray question marks. */
export function cleanName(s) {
  return String(s || '').replace(/\?+/g, '').replace(/\s+/g, ' ').trim();
}

/**
 * Builds a matcher for a list of roster names.
 * A roster name matches when its first name is followed - within a few words -
 * by its last name, so middle names may be dropped.
 */
export function createNameMatcher(names) {
  const entries = names
    .filter(Boolean)
    .map(name => ({ name, tokens: normalizeName(name).split(' ').filter(Boolean) }))
    .filter(e => e.tokens.length)
    .sort((a, b) => b.tokens.length - a.tokens.length);

  /** Every roster name in a text, in order, with its character range in the text. */
  function findAll(text) {
    const words = [];
    for (const m of String(text).matchAll(/[^\s\-–—_.,:;()+/]+/g)) {
      const word = normalizeName(m[0]);
      if (word) words.push({ word, from: m.index, to: m.index + m[0].length });
    }
    const found = [];
    const used = new Set();
    const free = (i, j) => { for (let k = i; k <= j; k++) if (used.has(k)) return false; return true; };
    for (const e of entries) {
      const first = e.tokens[0];
      const last = e.tokens[e.tokens.length - 1];
      for (let i = 0; i < words.length; i++) {
        if (used.has(i) || words[i].word !== first) continue;
        let end = e.tokens.length === 1 ? i : -1;
        for (let j = i + 1; end < 0 && j <= Math.min(words.length - 1, i + e.tokens.length + 1); j++) {
          if (words[j].word === last && free(i, j)) end = j;
        }
        if (end < 0) continue;
        for (let k = i; k <= end; k++) used.add(k);
        found.push({ name: e.name, start: i, from: words[i].from, to: words[end].to });
      }
    }
    return found.sort((a, b) => a.start - b.start);
  }

  /** The roster name for a short piece of text that should be a single name, or null. */
  function match(text) {
    const hits = findAll(text);
    return hits.length === 1 ? hits[0].name : null;
  }

  return { findAll, match };
}
