// Where the app keeps its data and how it hands out files.
// Inside a claude.ai artifact: the viewer's private db subtree and the downloads
// capability. Anywhere else (GitHub Pages, a local server): localStorage and a
// normal browser download.

const LOCAL_KEY = 'nsm-shibutz-v1';

const capability = name => (window.claude && typeof window.claude.use === 'function' ? window.claude.use(name).catch(() => null) : Promise.resolve(null));

/**
 * The state is split so each stored piece stays small: one document for the
 * instructors, settings and notes, and one per month of events.
 */
function splitState(state) {
  const { events = [], ...rest } = state;
  const months = {};
  for (const ev of events) (months[ev.date.slice(0, 7)] ||= []).push(ev);
  return { main: rest, months };
}

function joinState(main, months) {
  const events = Object.keys(months).sort().flatMap(k => months[k]);
  return { ...main, events };
}

async function openDbStore() {
  const [db, user] = await Promise.all([capability('db'), capability('user')]);
  if (!db || !user) return null;
  let uid = null;
  try { uid = await user.id(); } catch { uid = null; }
  if (!uid) return null;
  const root = db.doc(`data/users/${uid}/state`);
  const months = root.collection('months');
  const written = new Map(); // what is already stored, to write only what changed

  return {
    kind: 'db',
    async load() {
      const [mainSnap, monthSnap] = await Promise.all([root.get(), months.get()]);
      if (!mainSnap.exists) return null;
      const main = JSON.parse(mainSnap.data().json || '{}');
      written.set('main', mainSnap.data().json);
      const byMonth = {};
      for (const d of monthSnap.docs) {
        const json = d.data().json || '[]';
        byMonth[d.id] = JSON.parse(json);
        written.set(d.id, json);
      }
      return joinState(main, byMonth);
    },
    async save(state) {
      const { main, months: byMonth } = splitState(state);
      const mainJson = JSON.stringify(main);
      if (written.get('main') !== mainJson) {
        await root.set({ json: mainJson, savedAt: new Date().toISOString() });
        written.set('main', mainJson);
      }
      for (const [key, list] of Object.entries(byMonth)) {
        const json = JSON.stringify(list);
        if (written.get(key) === json) continue;
        await months.doc(key).set({ json });
        written.set(key, json);
      }
      for (const key of [...written.keys()]) {
        if (key !== 'main' && !(key in byMonth)) { await months.doc(key).delete(); written.delete(key); }
      }
    },
  };
}

function localStore() {
  return {
    kind: 'local',
    async load() {
      try {
        const raw = localStorage.getItem(LOCAL_KEY);
        return raw ? JSON.parse(raw) : null;
      } catch { return null; }
    },
    async save(state) {
      try { localStorage.setItem(LOCAL_KEY, JSON.stringify(state)); } catch (e) { throw new Error('השמירה בדפדפן נכשלה: ' + e.message); }
    },
  };
}

/** The best store this page can use, chosen once at startup. */
export async function openStore() {
  try {
    const db = await openDbStore();
    if (db) return db;
  } catch { /* fall back to the browser */ }
  return localStore();
}

/** Save a generated file: through the artifact's downloads capability, or a normal download. */
export async function saveFile(filename, data) {
  const downloads = await capability('downloads');
  if (downloads) {
    try {
      await downloads.save({ filename, data });
      return 'saved';
    } catch (e) {
      if (e && e.code === 'declined') return 'declined';
      if (e && !['unavailable', 'not_granted', 'capability_disabled', 'capability_removed'].includes(e.code)) throw e;
    }
  }
  const blob = data instanceof Blob ? data : new Blob([data]);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  return 'saved';
}

const loaded = new Map();
/** Load a UMD script once and resolve with the global it defines. */
export function loadScript(src, globalName) {
  if (window[globalName]) return Promise.resolve(window[globalName]);
  if (!loaded.has(src)) {
    loaded.set(src, new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => (window[globalName] ? resolve(window[globalName]) : reject(new Error(`${globalName} לא נטען`)));
      s.onerror = () => { loaded.delete(src); reject(new Error('טעינת ספרייה נכשלה - בדקו חיבור לאינטרנט')); };
      document.head.appendChild(s);
    }));
  }
  return loaded.get(src);
}
