import { rtdb } from './firebase-config.js';
import {
  ref, push, set
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js";

const ROOT = 'logs';
const MAX_META_TEXT = 500;

function _pad(n) {
  return String(n).padStart(2, '0');
}

function _localDateKey(date) {
  return `${date.getFullYear()}-${_pad(date.getMonth() + 1)}-${_pad(date.getDate())}`;
}

function _sanitizeMeta(value, depth = 0) {
  if (value === null || value === undefined) return value;
  if (depth > 3) return '[max-depth]';
  if (typeof value === 'string') return value.length > MAX_META_TEXT ? value.slice(0, MAX_META_TEXT) + '...' : value;
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  if (Array.isArray(value)) return value.slice(0, 20).map(item => _sanitizeMeta(item, depth + 1));
  if (typeof value !== 'object') return String(value);

  const out = {};
  Object.keys(value).slice(0, 30).forEach(key => {
    out[key.replace(/[.#$/\[\]\u0000-\u001f\u007f]/g, '_') || '_empty'] = _sanitizeMeta(value[key], depth + 1);
  });
  return out;
}

export async function writeLog(entry) {
  const now = new Date();
  const system = String(entry.system || 'unknown').replace(/[.#$/\[\]\u0000-\u001f\u007f]/g, '_');
  const level = entry.level || 'info';
  const dateKey = _localDateKey(now);
  const logRef = push(ref(rtdb, `${ROOT}/${system}/${dateKey}`));

  await set(logRef, {
    time: now.toISOString(),
    system,
    level,
    action: entry.action || '',
    message: entry.message || '',
    durationMs: entry.durationMs ?? null,
    source: entry.source || 'config.js',
    page: typeof location !== 'undefined' ? location.pathname + location.search : '',
    meta: _sanitizeMeta(entry.meta || {})
  });
}
