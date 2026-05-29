/* =========================================================
   Tainty — auth.js
   Sesión global compartida entre módulos. Reemplaza al PIN
   por-módulo. La sesión vive en sessionStorage, así que se
   mantiene mientras el navegador esté abierto y se borra al
   cerrarlo o llamar a logout().
   ---------------------------------------------------------
   Módulos protegidos llaman requireLogin() al inicio de su
   script para redirigir a index.html si no hay sesión.
   panel_produccion es público (no llama requireLogin) — solo
   sus funciones de administración consultan isLoggedIn().
   ========================================================= */

const TAINTY_PIN = '200720';
const SESSION_KEY = 'tainty_session';

/** True si hay sesión activa en este navegador. */
function isLoggedIn() {
  try { return sessionStorage.getItem(SESSION_KEY) === '1'; }
  catch (_) { return false; }
}

/** Valida el PIN. Devuelve true si correcto y crea la sesión. */
function login(pin) {
  if (pin === TAINTY_PIN) {
    try { sessionStorage.setItem(SESSION_KEY, '1'); } catch (_) {}
    return true;
  }
  return false;
}

/** Borra la sesión. */
function logout() {
  try { sessionStorage.removeItem(SESSION_KEY); } catch (_) {}
}

/**
 * Verifica que haya sesión; si no, redirige al index y devuelve false.
 * Llamar en el primer momento del script de cada módulo protegido.
 */
function requireLogin() {
  if (!isLoggedIn()) {
    window.location.href = './index.html';
    return false;
  }
  return true;
}

/** Botón "Salir": borra sesión y vuelve al index. */
function cerrarSesion() {
  logout();
  window.location.href = './index.html';
}
