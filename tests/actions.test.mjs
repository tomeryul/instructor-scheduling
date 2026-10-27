import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// Every button the page draws with data-act="…" must have code that handles it;
// a missing case makes the button silently do nothing.
test('every data-act button has a handler', () => {
  const src = readFileSync(new URL('../js/app.js', import.meta.url), 'utf8');
  const acts = [...new Set([...src.matchAll(/data-act="([a-z-]+)"/g)].map(m => m[1]))];
  const handled = act => src.includes(`case '${act}'`) || src.includes(`act === '${act}'`) || src.includes(`[data-act="${act}"]`);
  const missing = acts.filter(a => !handled(a));
  assert.deepEqual(missing, []);
  assert.ok(acts.length > 20);
});
