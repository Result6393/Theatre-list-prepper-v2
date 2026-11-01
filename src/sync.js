// Work settings sync — shared client module.
// Canonical copy: https://github.com/Result6393/Result6393.github.io/blob/main/settings/sync.js
// Each app keeps an identical copy (offline apps can't load it from elsewhere); copy changes across.
//
// How it works
// - Your password is stretched (PBKDF2-SHA256, 600k rounds) into two keys:
//     an AES-GCM key that encrypts settings in the browser, and
//     an auth token the Worker checks (it only stores the token's hash).
// - Both are kept in this browser under SYNC_KEY. All apps on result6393.github.io share
//   that storage, so signing in once covers every work app on this device.
// - Apps keep saving settings to their own localStorage key as before. This module:
//     pullBeforeStart(app, key)  — before the app starts, replaces the local copy if the
//                                  server's is newer (2.5 s timeout; offline just uses local)
//     pushSoon(app, key)         — after a local save, uploads it (debounced)
//     mountSyncControl(el, app, key, onApplied) — small sign-in / status UI

export const WORKER_URL = 'https://work-settings.aflibercept.workers.dev';
const SYNC_KEY = 'work-settings-sync';
const SALT = 'result6393-work-settings-v1';
const ITERATIONS = 600000;
const PULL_TIMEOUT_MS = 2500;

// ---------- small helpers ----------
const enc = new TextEncoder();
const dec = new TextDecoder();
const b64 = (bytes) => btoa(String.fromCharCode(...new Uint8Array(bytes)));
const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const hex = (bytes) => [...new Uint8Array(bytes)].map((b) => b.toString(16).padStart(2, '0')).join('');

function store() {
  try {
    return window.localStorage;
  } catch {
    return null;
  }
}
function readJson(key) {
  try {
    return JSON.parse(store()?.getItem(key) || 'null');
  } catch {
    return null;
  }
}
function writeJson(key, value) {
  try {
    if (value === null) store()?.removeItem(key);
    else store()?.setItem(key, JSON.stringify(value));
  } catch {
    /* storage blocked: sync just won't persist */
  }
}
const metaKey = (localKey) => `${localKey}.sync`;

// ---------- keys ----------
/** Derives { authToken (hex), encKey (base64 raw AES key) } from a password. */
export async function deriveKeys(password) {
  const base = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits(
    { name: 'PBKDF2', hash: 'SHA-256', salt: enc.encode(SALT), iterations: ITERATIONS },
    base,
    512,
  );
  const all = new Uint8Array(bits);
  return { authToken: hex(all.slice(32)), encKey: b64(all.slice(0, 32)) };
}

/** SHA-256 of the auth token: the value to store as the Worker's AUTH_HASH secret. */
export async function authHashFor(password) {
  const { authToken } = await deriveKeys(password);
  return hex(await crypto.subtle.digest('SHA-256', enc.encode(authToken)));
}

function creds() {
  const c = readJson(SYNC_KEY);
  return c && typeof c.authToken === 'string' && typeof c.encKey === 'string' ? c : null;
}
export const isSignedIn = () => creds() !== null;

async function aesKey(c) {
  return crypto.subtle.importKey('raw', unb64(c.encKey), 'AES-GCM', false, ['encrypt', 'decrypt']);
}

// ---------- server calls ----------
async function call(method, path, c, body, timeoutMs) {
  const ctrl = new AbortController();
  const t = timeoutMs ? setTimeout(() => ctrl.abort(), timeoutMs) : null;
  try {
    return await fetch(WORKER_URL + path, {
      method,
      headers: { Authorization: `Bearer ${c.authToken}`, ...(body ? { 'Content-Type': 'application/json' } : {}) },
      body: body ? JSON.stringify(body) : undefined,
      signal: ctrl.signal,
      cache: 'no-store',
    });
  } finally {
    if (t) clearTimeout(t);
  }
}

async function fetchRemote(app, c, timeoutMs) {
  const res = await call('GET', `/s/${app}`, c, null, timeoutMs);
  if (res.status === 404) return null;
  if (res.status === 401) throw Object.assign(new Error('Sync password no longer valid. Sign in again.'), { signedOut: true });
  if (!res.ok) throw new Error(`Sync server error ${res.status}`);
  const blob = await res.json();
  try {
    const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(blob.iv) }, await aesKey(c), unb64(blob.ct));
    return { text: dec.decode(plain), updated: blob.updated };
  } catch {
    // Saved under a previous password: treat as empty so this device's settings replace it.
    return null;
  }
}

async function upload(app, c, text, updated) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await aesKey(c), enc.encode(text));
  const res = await call('PUT', `/s/${app}`, c, { iv: b64(iv), ct: b64(ct), updated }, 10000);
  if (res.status === 401) throw Object.assign(new Error('Sync password no longer valid. Sign in again.'), { signedOut: true });
  if (!res.ok) throw new Error(`Sync server error ${res.status}`);
}

// ---------- public API ----------

/** Sign in on this device. Throws with a readable message on a wrong password or lockout. */
export async function signIn(password) {
  const c = await deriveKeys(password);
  const res = await fetch(WORKER_URL + '/auth', { method: 'POST', headers: { Authorization: `Bearer ${c.authToken}` } });
  if (res.status === 401) throw new Error('Wrong password.');
  if (res.status === 429) throw new Error('Too many wrong passwords. Try again in 15 minutes.');
  if (!res.ok) throw new Error(`Sync server error ${res.status}`);
  writeJson(SYNC_KEY, c);
}

export function signOut() {
  writeJson(SYNC_KEY, null);
}

/**
 * Call before the app reads its settings. Returns 'pulled' | 'pushed' | 'same' | 'offline' | 'signed-out'.
 * Newest copy wins: if this device saved while offline, it uploads instead.
 */
export async function pullBeforeStart(app, localKey, timeoutMs = PULL_TIMEOUT_MS) {
  const c = creds();
  if (!c) return 'signed-out';
  const meta = readJson(metaKey(localKey)) || { updated: 0, dirty: false };
  try {
    const remote = await fetchRemote(app, c, timeoutMs);
    const localText = store()?.getItem(localKey);
    if (remote && remote.updated > meta.updated) {
      store()?.setItem(localKey, remote.text);
      writeJson(metaKey(localKey), { updated: remote.updated, dirty: false });
      return 'pulled';
    }
    if (localText && (meta.dirty || !remote)) {
      const updated = meta.updated || Date.now();
      await upload(app, c, localText, updated);
      writeJson(metaKey(localKey), { updated, dirty: false });
      return 'pushed';
    }
    return 'same';
  } catch (e) {
    if (e && e.signedOut) signOut();
    return 'offline';
  }
}

const timers = {};
/** Call after the app saves its settings to localStorage. Uploads a second later. */
export function pushSoon(app, localKey, delayMs = 1000) {
  const updated = Date.now();
  writeJson(metaKey(localKey), { updated, dirty: true });
  const c = creds();
  if (!c) return;
  clearTimeout(timers[localKey]);
  timers[localKey] = setTimeout(async () => {
    const text = store()?.getItem(localKey);
    if (text == null) return;
    try {
      const m = readJson(metaKey(localKey)) || { updated };
      await upload(app, c, text, m.updated);
      writeJson(metaKey(localKey), { updated: m.updated, dirty: false });
      notify();
    } catch (e) {
      if (e && e.signedOut) signOut();
      notify();
      // stays dirty; next start-up retries
    }
  }, delayMs);
}

const listeners = new Set();
const notify = () => listeners.forEach((f) => f());

/**
 * Renders a small sync panel into `el`.
 * onApplied() is called after sign-in brings down settings from another device; the app
 * should reload its settings (simplest: location.reload()).
 */
export function mountSyncControl(el, app, localKey, onApplied) {
  const root = document.createElement('div');
  root.className = 'work-sync';
  root.style.cssText = 'border:1px solid rgba(127,127,127,.35);border-radius:8px;padding:10px 12px;margin:8px 0;font-size:14px;line-height:1.4';
  el.appendChild(root);

  const render = (msg = '') => {
    root.textContent = '';
    const title = document.createElement('div');
    title.style.fontWeight = '600';
    title.textContent = 'Sync settings across devices';
    root.appendChild(title);

    if (isSignedIn()) {
      const meta = readJson(metaKey(localKey));
      const status = document.createElement('div');
      status.textContent = meta?.dirty
        ? 'Signed in. Latest change not uploaded yet (offline?). It will retry.'
        : 'Signed in. Changes here sync to your other devices.';
      root.appendChild(status);
      const out = button('Sign out on this device', () => {
        signOut();
        render('Signed out. Settings stay on this device.');
      });
      root.appendChild(out);
    } else {
      const form = document.createElement('form');
      form.style.cssText = 'display:flex;gap:6px;flex-wrap:wrap;margin-top:6px';
      const pw = document.createElement('input');
      pw.type = 'password';
      pw.placeholder = 'Sync password';
      pw.autocomplete = 'current-password';
      pw.style.cssText = 'flex:1;min-width:160px;padding:6px 8px;font:inherit';
      const go = document.createElement('button');
      go.type = 'submit';
      go.textContent = 'Sign in';
      form.append(pw, go);
      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        go.disabled = true;
        go.textContent = 'Signing in…';
        try {
          await signIn(pw.value);
          const result = await pullBeforeStart(app, localKey, 10000);
          if (result === 'pulled') {
            render('Signed in. Settings from your other devices were downloaded.');
            onApplied && onApplied();
          } else {
            render(result === 'pushed' ? 'Signed in. This device’s settings were uploaded.' : 'Signed in.');
          }
        } catch (err) {
          render(err && err.message ? err.message : 'Could not reach the sync server.');
        }
      });
      root.appendChild(form);
    }
    if (msg) {
      const m = document.createElement('div');
      m.style.cssText = 'margin-top:6px;opacity:.8';
      m.textContent = msg;
      root.appendChild(m);
    }
  };

  function button(label, onClick) {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = label;
    b.style.marginTop = '6px';
    b.addEventListener('click', onClick);
    return b;
  }

  const rerender = () => render();
  listeners.add(rerender);
  render();
  return () => {
    listeners.delete(rerender);
    root.remove();
  };
}
