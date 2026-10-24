// One JSON file in a private GitHub repository, read and written with the
// user's own fine-grained token. GitHub keeps every version as a commit; the
// file's sha tells whether another device wrote it in between.

const API = 'https://api.github.com';
export const DEFAULT_PATH = 'shibutz.json';

export class SyncError extends Error {
  constructor(message, code) { super(message); this.code = code; }
}

/** "owner/repo", or a github.com link to it. */
export function parseRepo(text) {
  const m = String(text || '').trim().replace(/\.git$/, '').match(/(?:github\.com\/)?([\w.-]+)\/([\w.-]+)\/?$/);
  return m ? { owner: m[1], repo: m[2] } : null;
}

function utf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

function base64ToUtf8(b64) {
  const bin = atob(b64.replace(/\s/g, ''));
  return new TextDecoder().decode(Uint8Array.from(bin, c => c.charCodeAt(0)));
}

async function call(cfg, path, { method = 'GET', body, accept = 'application/vnd.github+json', etag } = {}) {
  const headers = { Accept: accept, Authorization: `Bearer ${cfg.token}`, 'X-GitHub-Api-Version': '2022-11-28' };
  if (body) headers['Content-Type'] = 'application/json';
  if (etag) headers['If-None-Match'] = etag;
  let res;
  try {
    res = await fetch(API + path, { method, headers, body: body && JSON.stringify(body), cache: 'no-store' });
  } catch {
    throw new SyncError('אין חיבור לאינטרנט', 'offline');
  }
  if (res.status === 401) throw new SyncError('הטוקן לא תקין או שפג תוקפו', 'auth');
  if (res.status === 403 && res.headers.get('x-ratelimit-remaining') === '0') throw new SyncError('GitHub מגביל זמנית את מספר הבקשות - ננסה שוב בעוד כמה דקות', 'rate');
  if (res.status === 403) throw new SyncError('לטוקן אין הרשאת כתיבה לקבצים בריפו (Contents: Read and write)', 'forbidden');
  return res;
}

const repoPath = cfg => `/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}`;
const filePath = cfg => `${repoPath(cfg)}/contents/${(cfg.path || DEFAULT_PATH).split('/').map(encodeURIComponent).join('/')}`;

/** Check the token reaches the repository and that it is private. */
export async function checkRepo(cfg) {
  const res = await call(cfg, repoPath(cfg));
  if (res.status === 404) throw new SyncError('הריפו לא נמצא, או שלטוקן אין גישה אליו', 'missing');
  if (!res.ok) throw new SyncError(`GitHub החזיר שגיאה ${res.status}`, 'http');
  const repo = await res.json();
  if (!repo.private) throw new SyncError('הריפו הזה ציבורי, וכל אחד היה רואה בו את פרטי המדריכים. לנתונים צריך ריפו נוסף ונפרד, פרטי (הריפו של האתר נשאר ציבורי בשביל GitHub Pages).', 'public');
  return repo;
}

/**
 * The file and its sha, null when it does not exist yet, or
 * { notModified: true } when it has not changed since `etag`.
 */
export async function readFile(cfg, etag) {
  const res = await call(cfg, filePath(cfg), { etag });
  if (res.status === 304) return { notModified: true };
  if (res.status === 404) return null;
  if (!res.ok) throw new SyncError(`קריאת הקובץ מ-GitHub נכשלה (${res.status})`, 'http');
  const meta = await res.json();
  let text;
  if (meta.content && meta.encoding === 'base64') text = base64ToUtf8(meta.content);
  else {
    // Files over 1 MB come without content in this call.
    const raw = await call(cfg, filePath(cfg), { accept: 'application/vnd.github.raw+json' });
    if (!raw.ok) throw new SyncError(`קריאת הקובץ מ-GitHub נכשלה (${raw.status})`, 'http');
    text = await raw.text();
  }
  let data;
  try { data = JSON.parse(text); } catch { throw new SyncError('הקובץ ב-GitHub פגום (לא JSON)', 'corrupt'); }
  return { sha: meta.sha, data, etag: res.headers.get('etag') };
}

/** Write the file; `sha` is the version it replaces. Throws code "conflict" if someone wrote it since. */
export async function writeFile(cfg, text, sha, message) {
  const res = await call(cfg, filePath(cfg), {
    method: 'PUT',
    body: { message, content: utf8ToBase64(text), ...(sha ? { sha } : {}) },
  });
  if (res.status === 409 || (res.status === 422 && sha)) throw new SyncError('הקובץ השתנה בינתיים', 'conflict');
  if (res.status === 404) throw new SyncError('הריפו לא נמצא, או שלטוקן אין גישה אליו', 'missing');
  if (!res.ok) throw new SyncError(`השמירה ב-GitHub נכשלה (${res.status})`, 'http');
  const out = await res.json();
  return out.content.sha;
}

/** The link that connects another device: the repository and the token, in the URL fragment only. */
export function connectLink(cfg, base) {
  const payload = JSON.stringify({ r: `${cfg.owner}/${cfg.repo}`, p: cfg.path || DEFAULT_PATH, t: cfg.token });
  const b64 = utf8ToBase64(payload).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  return `${base}#connect=${b64}`;
}

export function readConnectLink(hash) {
  const m = String(hash || '').match(/connect=([\w-]+)/);
  if (!m) return null;
  try {
    const b64 = m[1].replace(/-/g, '+').replace(/_/g, '/');
    const { r, p, t } = JSON.parse(base64ToUtf8(b64 + '='.repeat((4 - (b64.length % 4)) % 4)));
    const repo = parseRepo(r);
    return repo && t ? { ...repo, path: p || DEFAULT_PATH, token: t } : null;
  } catch { return null; }
}
