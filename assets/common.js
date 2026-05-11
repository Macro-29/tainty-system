/* =========================================================
   Tainty — common.js
   Shared client-side utilities.
   Requires: config.js loaded BEFORE this file (for SCRIPT_URL).
   ========================================================= */

/**
 * JSONP helper for Google Apps Script GET requests.
 * Returns a Promise that resolves with the parsed JSON response.
 * @param {string} url - full URL (will append callback= automatically)
 * @param {number} [timeoutMs=15000] - timeout in ms
 */
function jsonp(url, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const cb = 'cb_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    const s = document.createElement('script');
    const t = setTimeout(() => { cleanup(); reject(new Error('timeout')); }, timeoutMs);
    function cleanup() {
      clearTimeout(t);
      delete window[cb];
      if (s.parentNode) s.parentNode.removeChild(s);
    }
    window[cb] = data => { cleanup(); resolve(data); };
    s.src = url + (url.includes('?') ? '&' : '?') + 'callback=' + cb;
    s.onerror = () => { cleanup(); reject(new Error('error')); };
    document.head.appendChild(s);
  });
}

/**
 * POST helper for Google Apps Script using FormData + no-cors.
 * Note: response is opaque (no-cors), so we cannot read it back.
 * Wait a bit after the call to let the Sheet update before re-fetching.
 * @param {string} url - Apps Script URL
 * @param {Object} data - key/value pairs sent as FormData
 */
async function postToScript(url, data) {
  const form = new FormData();
  Object.entries(data).forEach(([k, v]) => form.append(k, v));
  return fetch(url, { method: 'POST', body: form, mode: 'no-cors' });
}

/**
 * Show a toast notification. Auto-creates the #toast element if absent.
 * @param {string} msg
 * @param {number|Object} [optsOrDuration=2500] - ms to auto-hide, or options object.
 *   Pass 0 to keep persistent (manual ocultarToast() to dismiss).
 *   Pass `{duration, type}` for full control. type: 'error' | 'success' | 'info'.
 */
function mostrarToast(msg, optsOrDuration = 2500) {
  const opts = typeof optsOrDuration === 'number'
    ? { duration: optsOrDuration }
    : (optsOrDuration || {});
  let t = document.getElementById('toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'toast';
    t.className = 'toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.classList.remove('error', 'success');
  if (opts.type === 'error') t.classList.add('error');
  else if (opts.type === 'success') t.classList.add('success');
  t.classList.add('visible');
  clearTimeout(t._hideTimer);
  if (opts.duration && opts.duration > 0) {
    t._hideTimer = setTimeout(() => t.classList.remove('visible'), opts.duration);
  }
}

/** Hide the toast immediately. */
function ocultarToast() {
  const t = document.getElementById('toast');
  if (t) {
    clearTimeout(t._hideTimer);
    t.classList.remove('visible');
  }
}

/**
 * Show/hide a generic loading overlay. Auto-creates if missing.
 * Usage: mostrarLoading('Cargando datos...'); ... ocultarLoading();
 * If the page already has a `<div id="loading"></div>` with class .loading-overlay,
 * this just toggles the .hidden class.
 */
function mostrarLoading(msg = 'Cargando...') {
  let el = document.getElementById('loading');
  if (!el) {
    el = document.createElement('div');
    el.id = 'loading';
    el.className = 'loading-overlay';
    el.innerHTML = '<div class="spinner"></div><div class="loading-text"></div>';
    document.body.appendChild(el);
  }
  const txt = el.querySelector('.loading-text');
  if (txt) txt.textContent = msg;
  el.classList.remove('hidden');
}

function ocultarLoading() {
  const el = document.getElementById('loading');
  if (el) el.classList.add('hidden');
}

/**
 * Wrap an async call (jsonp or fetch) with a standardized error toast.
 * Throws the error after toasting, so callers can still react via try/catch.
 * @param {Promise<any>} promise
 * @param {string} [contextMsg='Error al conectar'] - prefix for the toast
 */
async function conToastError(promise, contextMsg = 'Error al conectar') {
  try {
    return await promise;
  } catch (e) {
    const detail = e && e.message ? e.message : 'desconocido';
    mostrarToast(`${contextMsg}: ${detail}`, { type: 'error', duration: 3500 });
    throw e;
  }
}

/**
 * Infer product type from its name (looks for "(...cap)", "(...gr)" → defaults to "ml").
 * @param {string} nombre
 * @returns {'cap'|'gr'|'ml'}
 */
function getTipo(nombre) {
  const n = (nombre || '').toLowerCase();
  if (/\(.*cap\)/.test(n)) return 'cap';
  if (/\(.*gr\)/.test(n)) return 'gr';
  return 'ml';
}

/**
 * Format a number as Peruvian Sol currency. Example: 1234.5 → "S/ 1,234.50".
 */
function formatSoles(n) {
  const num = Number(n) || 0;
  return 'S/ ' + num.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Format a date as DD/MM (short) or DD/MM/YYYY (long).
 * @param {string|Date} d
 * @param {boolean} [withYear=false]
 */
function formatFecha(d, withYear = false) {
  const date = d instanceof Date ? d : new Date(d);
  if (isNaN(date)) return '';
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  return withYear ? `${dd}/${mm}/${date.getFullYear()}` : `${dd}/${mm}`;
}

// Note: individual pages declare their own MESES_ES (some uppercase, some title case).
// Keeping per-page to avoid const redeclaration conflicts.
