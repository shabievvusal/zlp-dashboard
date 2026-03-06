/**
 * auth.js — авторизация, хранение токена, автообновление
 *
 * Стратегия сессии:
 *   - accessToken хранится только в памяти (короткоживущий, ~5 мин)
 *   - refreshToken сохраняется в localStorage — переживает перезагрузку страницы
 *   - При старте: берём refreshToken из localStorage и делаем обновление
 */

import * as api from './api.js';

const LS_REFRESH_KEY = 'wms_refresh_token';

let accessToken = null;
let accessTokenExpiry = null;
let refreshToken = null;
let refreshTimer = null;
let onAuthChange = null;

export function getToken() {
  return accessToken;
}

export function isLoggedIn() {
  return !!accessToken;
}

export function setOnAuthChange(cb) {
  onAuthChange = cb;
}

function notifyChange(loggedIn) {
  if (onAuthChange) onAuthChange(loggedIn);
}

// ─── Сохранение refreshToken в localStorage ──────────────────────────────────

function saveRefreshToken(token) {
  refreshToken = token;
  if (token) {
    try { localStorage.setItem(LS_REFRESH_KEY, token); } catch { /* ignore */ }
  } else {
    try { localStorage.removeItem(LS_REFRESH_KEY); } catch { /* ignore */ }
  }
}

function loadRefreshTokenFromStorage() {
  try { return localStorage.getItem(LS_REFRESH_KEY) || null; } catch { return null; }
}

// ─── Публичные функции ───────────────────────────────────────────────────────

export async function login(loginValue, password) {
  const data = await api.loginSamokat(loginValue, password);
  if (!data?.value?.accessToken) throw new Error('Токен не получен в ответе');

  accessToken = data.value.accessToken;
  accessTokenExpiry = Date.now() + (data.value.expiresIn || 300) * 1000;
  saveRefreshToken(data.value.refreshToken || null);

  // Сохраняем токены на сервер (для автосбора в фоне)
  await api.putConfig({ token: accessToken, refreshToken: refreshToken || '' });

  scheduleRefresh();
  notifyChange(true);
}

export async function logout() {
  accessToken = null;
  accessTokenExpiry = null;
  saveRefreshToken(null);
  if (refreshTimer) { clearTimeout(refreshTimer); refreshTimer = null; }
  notifyChange(false);
}

/**
 * Пытается восстановить сессию из localStorage при загрузке страницы.
 * Возвращает true если успешно.
 */
export async function tryRestoreSession() {
  const stored = loadRefreshTokenFromStorage();
  if (!stored) {
    notifyChange(false);
    return false;
  }
  refreshToken = stored;
  const ok = await doRefresh();
  if (!ok) {
    saveRefreshToken(null); // невалидный — чистим
    notifyChange(false);
  }
  return ok;
}

// ─── Внутренние функции ──────────────────────────────────────────────────────

async function doRefresh() {
  if (!refreshToken) return false;
  try {
    const data = await api.refreshSamokatToken(refreshToken);
    if (!data?.value?.accessToken) return false;

    accessToken = data.value.accessToken;
    accessTokenExpiry = Date.now() + (data.value.expiresIn || 300) * 1000;

    // Обновляем refreshToken если пришёл новый
    if (data.value.refreshToken) saveRefreshToken(data.value.refreshToken);

    // Синхронизируем с сервером для фонового автосбора
    await api.putConfig({ token: accessToken, refreshToken: refreshToken || '' });

    notifyChange(true);
    scheduleRefresh();
    return true;
  } catch {
    return false;
  }
}

function scheduleRefresh() {
  if (refreshTimer) clearTimeout(refreshTimer);
  if (!accessTokenExpiry) return;
  // Обновляем за 60 секунд до истечения, минимум через 10 секунд
  const delay = Math.max(10000, accessTokenExpiry - Date.now() - 60000);
  refreshTimer = setTimeout(async () => {
    const ok = await doRefresh();
    if (!ok) {
      accessToken = null;
      saveRefreshToken(null);
      notifyChange(false);
    }
  }, delay);
}
