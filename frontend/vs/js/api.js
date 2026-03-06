/**
 * api.js — все запросы к backend API
 */

const API = '/api';

export async function getStatus() {
  const r = await fetch(`${API}/status`);
  return r.json();
}

export async function getConfig() {
  const r = await fetch(`${API}/config`);
  return r.json();
}

export async function putConfig(data) {
  const r = await fetch(`${API}/config`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return r.json();
}

export async function fetchData(options = {}, token) {
  const body = { options };
  if (token) body.token = token;
  const r = await fetch(`${API}/fetch-data`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  return r.json();
}

export async function listShifts() {
  const r = await fetch(`${API}/shifts`);
  return r.json();
}

export async function getCurrentShift() {
  const r = await fetch(`${API}/shifts/current`);
  return r.json();
}

export async function getShiftItems(shiftKey) {
  const r = await fetch(`${API}/shifts/${encodeURIComponent(shiftKey)}/items`);
  return r.json();
}

/** Операции за один календарный день. shift=day|night — только нужная смена (в разы меньше данных и быстрее). */
export async function getDateItems(date, { fromHour, toHour, shift } = {}) {
  const params = new URLSearchParams();
  if (shift === 'day' || shift === 'night') params.set('shift', shift);
  if (fromHour != null) params.set('fromHour', String(fromHour));
  if (toHour != null) params.set('toHour', String(toHour));
  const qs = params.toString();
  const url = `${API}/date/${encodeURIComponent(date)}/items` + (qs ? `?${qs}` : '');
  const r = await fetch(url);
  return r.json();
}

export async function scheduleStart() {
  const r = await fetch(`${API}/schedule/start`, { method: 'POST' });
  return r.json();
}

export async function scheduleStop() {
  const r = await fetch(`${API}/schedule/stop`, { method: 'POST' });
  return r.json();
}

export async function scheduleSettings(data) {
  // data: { intervalMinutes?, pageSize? }
  const r = await fetch(`${API}/schedule/settings`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return r.json();
}

export async function getEmployees() {
  const r = await fetch(`${API}/employees`);
  return r.json();
}

/** Отправить PNG по компаниям как файлы (документы) в Telegram. items: [{ blob, caption, filename }] */
export async function sendHourlyStatsTelegram(items) {
  const fd = new FormData();
  const captions = [];
  items.forEach((item, i) => {
    fd.append('documents', item.blob, item.filename || `hourly_${i + 1}.png`);
    captions.push(item.caption || '');
  });
  fd.append('captions', JSON.stringify(captions));
  const r = await fetch(`${API}/stats/send-hourly-telegram`, { method: 'POST', body: fd });
  return r.json();
}

/** Добавить/дописать одного сотрудника в empl.csv (как в настройках дашборда). */
export async function saveEmplOne(fio, company) {
  const r = await fetch(`${API}/empl`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fio: (fio || '').trim(), company: (company != null ? String(company) : '').trim() }),
  });
  return r.json();
}

const LIVE_MONITOR_URL = 'https://api.samokat.ru/wmsops-wwh/activity-monitor/selection/handling-units-in-progress';

/** Живой мониторинг через backend (токен с сервера). */
export async function getLiveMonitor() {
  const r = await fetch(`${API}/monitor/live`);
  const text = await r.text();
  let data;
  try { data = text ? JSON.parse(text) : null; } catch {
    throw new Error('Ответ не JSON: ' + (text || '').slice(0, 150));
  }
  if (!r.ok) {
    const msg = data?.error || r.statusText;
    throw new Error(msg || `HTTP ${r.status}`);
  }
  return data;
}

/** Живой мониторинг запросом из браузера (с токеном пользователя) — чтобы Samokat видел сессию. */
export async function getLiveMonitorViaBrowser(token) {
  const r = await fetch(LIVE_MONITOR_URL, {
    headers: {
      'Accept': 'application/json',
      'Authorization': `Bearer ${token}`,
      'Origin': 'https://wwh.samokat.ru',
      'Referer': 'https://wwh.samokat.ru/',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36',
    },
  });
  const text = await r.text();
  const trimmed = (text || '').trim().toLowerCase();
  if (trimmed.startsWith('<!doctype') || trimmed.startsWith('<html')) {
    throw new Error('Сервер вернул HTML вместо JSON. Проверьте вход или обновите страницу.');
  }
  let data;
  try { data = text ? JSON.parse(text) : null; } catch {
    throw new Error('Ответ не JSON: ' + (text || '').slice(0, 150));
  }
  if (!r.ok) {
    const msg = data?.message || data?.error || r.statusText;
    throw new Error(`API ${r.status}: ${msg}`);
  }
  return data;
}

export async function getRollcall() {
  const r = await fetch(`${API}/rollcall`);
  return r.json();
}

export async function putRollcall(shiftKey, present) {
  const r = await fetch(`${API}/rollcall`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ shiftKey, present }),
  });
  return r.json();
}

export async function loginSamokat(login, password) {
  const r = await fetch('https://api.samokat.ru/wmsin-wwh/auth/password', {
    method: 'POST',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'Origin': 'https://wwh.samokat.ru',
      'Referer': 'https://wwh.samokat.ru/',
    },
    body: JSON.stringify({ login, password }),
  });
  if (!r.ok) throw new Error(`Ошибка авторизации: ${r.status}`);
  return r.json();
}

export async function refreshSamokatToken(refreshToken) {
  const r = await fetch('https://api.samokat.ru/wmsin-wwh/auth/refresh', {
    method: 'POST',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'Origin': 'https://wwh.samokat.ru',
      'Referer': 'https://wwh.samokat.ru/',
    },
    body: JSON.stringify({ refreshToken }),
  });
  if (!r.ok) throw new Error(`Ошибка обновления токена: ${r.status}`);
  return r.json();
}

// ─── Консолидация ────────────────────────────────────────────────────────────

export async function getConsolidationComplaints() {
  const r = await fetch(`${API}/consolidation/complaints`);
  return r.json();
}

export async function updateComplaintStatus(id, status) {
  const r = await fetch(`${API}/consolidation/complaints/${encodeURIComponent(id)}/status`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ status }),
  });
  return r.json();
}

export async function deleteComplaint(id) {
  const r = await fetch(`${API}/consolidation/complaints/${encodeURIComponent(id)}`, {
    method: 'DELETE',
  });
  return r.json();
}

export async function saveComplaintLookup(id, data) {
  const r = await fetch(`${API}/consolidation/complaints/${encodeURIComponent(id)}/lookup`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return r.json();
}

export async function sendComplaintsToTelegram(complaintIds) {
  const r = await fetch(`${API}/consolidation/telegram/send`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ complaintIds }),
  });
  return r.json();
}

// ─── Подключение через браузер (страница /vs) ───────────────────────────────

const SAMOKAT_STOCKS_URL = 'https://api.samokat.ru/wmsops-wwh/stocks/changes/search';

function buildBodyForBrowser(options = {}) {
  let from = options.operationCompletedAtFrom;
  let to = options.operationCompletedAtTo;
  if (options.date) {
    const dateStr = String(options.date).slice(0, 10);
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      from = `${dateStr}T00:00:00.000Z`;
      to = `${dateStr}T23:59:59.999Z`;
    }
  }
  if (!from || !to) {
    const now = new Date();
    const h = now.getHours();
    const y = now.getFullYear();
    const m = now.getMonth();
    const d = now.getDate();
    if (h >= 9 && h < 21) {
      const fromDate = new Date(y, m, d, 9, 0, 0, 0);
      const toDate = new Date(y, m, d, 20, 59, 59, 999);
      from = fromDate.toISOString();
      to = toDate.toISOString();
    } else {
      const start = new Date(y, m, d, 21, 0, 0, 0);
      if (h < 9) start.setDate(start.getDate() - 1);
      const end = new Date(start);
      end.setDate(end.getDate() + 1);
      end.setHours(8, 59, 59, 999);
      from = start.toISOString();
      to = end.toISOString();
    }
  }
  return {
    productId: null,
    parts: [],
    operationTypes: ['PIECE_SELECTION_PICKING', 'PICK_BY_LINE'],
    sourceCellId: null,
    targetCellId: null,
    sourceHandlingUnitBarcode: null,
    targetHandlingUnitBarcode: null,
    operationStartedAtFrom: null,
    operationStartedAtTo: null,
    operationCompletedAtFrom: from,
    operationCompletedAtTo: to,
    executorId: null,
    pageNumber: options.pageNumber || 1,
    pageSize: options.pageSize || 2000,
  };
}

async function fetchOnePageFromBrowser(token, body) {
  const r = await fetch(SAMOKAT_STOCKS_URL, {
    method: 'POST',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'Origin': 'https://wwh.samokat.ru',
      'Referer': 'https://wwh.samokat.ru/',
      'Authorization': `Bearer ${token}`,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36',
    },
    body: JSON.stringify(body),
  });
  const text = await r.text();
  const trimmed = (text || '').trim().toLowerCase();
  if (trimmed.startsWith('<!doctype') || trimmed.startsWith('<html')) {
    throw new Error('Сервер вернул HTML вместо JSON. Проверьте VPN или доступ.');
  }
  let data;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    throw new Error('Ответ не JSON: ' + (text || '').slice(0, 150));
  }
  if (!r.ok) {
    const msg = data?.message || data?.error || r.statusText;
    throw new Error(`API ${r.status}: ${msg}`);
  }
  const value = data?.value || data;
  const items = Array.isArray(value?.items) ? value.items : [];
  const total = value?.total ?? data?.totalElements ?? null;
  return { items, total };
}

/** Точечный запрос: последняя активность по executorId за интервал (для мониторинга — 30 мин или 1 ч). */
export async function fetchLastCompletedForExecutor(token, executorId, fromIso, toIso) {
  const body = {
    productId: null,
    parts: [],
    operationTypes: ['PIECE_SELECTION_PICKING', 'PICK_BY_LINE'],
    sourceCellId: null,
    targetCellId: null,
    sourceHandlingUnitBarcode: null,
    targetHandlingUnitBarcode: null,
    operationStartedAtFrom: null,
    operationStartedAtTo: null,
    operationCompletedAtFrom: fromIso,
    operationCompletedAtTo: toIso,
    executorId: executorId || null,
    pageNumber: 1,
    pageSize: 100,
  };
  const { items } = await fetchOnePageFromBrowser(token, body);
  let maxCompletedAt = null;
  for (const item of items) {
    const at = item.operationCompletedAt;
    if (!at) continue;
    const ts = new Date(at).getTime();
    if (maxCompletedAt === null || ts > maxCompletedAt) maxCompletedAt = ts;
  }
  return { items, maxCompletedAt };
}

/** Группирует операции по (дата, час) по operationCompletedAt — время подтверждения/выполнения задачи. */
function groupItemsByHour(items) {
  const byHour = new Map();
  for (const item of items) {
    const ts = item.operationCompletedAt;
    if (!ts) continue;
    const d = new Date(ts);
    const dateStr = d.toISOString().slice(0, 10);
    const hour = d.getHours();
    const key = `${dateStr}\t${hour}`;
    if (!byHour.has(key)) byHour.set(key, []);
    byHour.get(key).push(item);
  }
  return byHour;
}

/** Загрузка данных через браузер (все страницы) и сохранение на сервер почасовыми порциями — много маленьких запросов вместо одного большого. */
export async function fetchDataViaBrowser(token, options = {}) {
  const pageSize = Math.min(2000, Math.max(100, parseInt(options.pageSize, 10) || 2000));
  const body = buildBodyForBrowser({ ...options, pageNumber: 1, pageSize });
  const first = await fetchOnePageFromBrowser(token, body);
  let allItems = [...first.items];
  let total = first.total ?? allItems.length;
  const totalPages = Math.ceil(total / pageSize);

  for (let p = 2; p <= totalPages; p++) {
    const nextBody = buildBodyForBrowser({ ...options, pageNumber: p, pageSize });
    const next = await fetchOnePageFromBrowser(token, nextBody);
    allItems = allItems.concat(next.items);
  }

  const byHour = groupItemsByHour(allItems);
  let totalAdded = 0;
  let totalSkipped = 0;
  const savedParts = [];

  for (const [key, items] of byHour) {
    const [dateStr, hour] = key.split('\t');
    const saveRes = await fetch(`${API}/save-fetched-data`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ value: { items, total: items.length } }),
    });
    const saveData = await saveRes.json();
    if (saveData.ok !== true) throw new Error(saveData.error || 'Ошибка сохранения');
    totalAdded += saveData.added ?? 0;
    totalSkipped += saveData.skipped ?? 0;
    savedParts.push(`${dateStr}-${hour}h`);
  }

  return {
    success: true,
    fetched: allItems.length,
    added: totalAdded,
    skipped: totalSkipped,
    savedTo: savedParts.length ? savedParts.join(', ') : '',
    total,
  };
}

/** Только сохранить уже полученные данные (например после ручной загрузки). */
export async function saveFetchedData(items) {
  const r = await fetch(`${API}/save-fetched-data`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ value: { items } }),
  });
  return r.json();
}
