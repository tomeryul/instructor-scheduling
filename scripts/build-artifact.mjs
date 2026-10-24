// Builds dist/artifact.html: the page content without the <html>/<head>/<body>
// wrapper, for publishing as a claude.ai artifact (which adds its own wrapper).
// The css/, js/ and vendor/ files are published next to it unchanged.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const head = html.match(/<head>([\s\S]*?)<\/head>/)[1]
  .replace(/<meta charset[^>]*>\s*/, '')
  .replace(/<meta name="viewport"[^>]*>\s*/, '');
const body = html.match(/<body>([\s\S]*?)<\/body>/)[1];
mkdirSync(new URL('../dist/', import.meta.url), { recursive: true });
writeFileSync(new URL('../dist/artifact.html', import.meta.url), `${head.trim()}\n${body.trim()}\n`);
console.log('dist/artifact.html written');
