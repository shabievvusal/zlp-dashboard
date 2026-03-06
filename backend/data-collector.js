const fs = require('fs');
const path = require('path');
const { loadConfig, ensureDataDir, DATA_DIR } = require('./scheduler');
const { buildDashboardSummary } = require('./build-dashboard-summary');

const API_URL = 'https://api.samokat.ru/wmsops-wwh/stocks/changes/search';

/** Таймаут запроса к API (мс). Если ответа нет — вернём ошибку вместо зависания. */
const FETCH_TIMEOUT_MS = 45000;

const DEFAULT_HEADERS = {
  'Accept': 'application/json',
  'Accept-Encoding': 'gzip, deflate, br',
  'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
  'Content-Type': 'application/json',
  'Origin': 'https://wwh.samokat.ru',
  'Referer': 'https://wwh.samokat.ru/',
  'Sec-Fetch-Dest': 'empty',
  'Sec-Fetch-Mode': 'cors',
  'Sec-Fetch-Site': 'same-site',
  'sec-ch-ua-platform': '"Windows"',
  'sec-ch-ua': '"Not(A:Brand";v="8", "Chromium";v="144", "Google Chrome";v="144"',
  'sec-ch-ua-mobile': '?0',
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36',
};

function buildBody(options = {}) {
  const from = options.operationCompletedAtFrom || new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString().slice(0, 10) + 'T21:00:00.000Z';
  const to = options.operationCompletedAtTo || new Date().toISOString().slice(0, 10) + 'T20:59:59.000Z';
  const operationTypes = Array.isArray(options.operationTypes) && options.operationTypes.length
    ? options.operationTypes
    : ['PIECE_SELECTION_PICKING', 'PICK_BY_LINE'];
  return {
    productId: null,
    parts: [],
    operationTypes,
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
    pageSize: options.pageSize || 100,
  };
}

const pad2 = (n) => String(n).padStart(2, '0');

/** Локальное время HH:mm из Date (для имени файла используем дефис: HH-mm, т.к. двоеточие не везде допустимо в имени файла). */
function toLocalTimePart(d) {
  return `${pad2(d.getHours())}-${pad2(d.getMinutes())}`;
}

/** Локальная дата DD.MM.YY */
function toLocalDateStr(d) {
  return [pad2(d.getDate()), pad2(d.getMonth() + 1), String(d.getFullYear()).slice(-2)].join('.');
}

/** Имя файла по диапазону запроса (запасной вариант): "09-00 - 17-13 26.02.26.json" (всё по локальному времени). */
function filenameFromRange(operationCompletedAtFrom, operationCompletedAtTo) {
  const fromDate = new Date(operationCompletedAtFrom);
  const toDate = new Date(operationCompletedAtTo);
  const fromPart = toLocalTimePart(fromDate);
  const toPart = toLocalTimePart(toDate);
  const dateStr = toLocalDateStr(toDate);
  return `${fromPart} - ${toPart} ${dateStr}.json`;
}

/** Имя файла по времени первой и последней записи: первый пик 9:02, последний 17:52 → "09-02 - 17-52 26.02.26.json" (локальное время). */
function filenameFromItems(items) {
  if (!Array.isArray(items) || items.length === 0) return null;
  const times = items
    .map((it) => it && it.operationCompletedAt)
    .filter(Boolean)
    .map((s) => new Date(s).getTime());
  if (times.length === 0) return null;
  const minT = Math.min(...times);
  const maxT = Math.max(...times);
  const fromDate = new Date(minT);
  const toDate = new Date(maxT);
  const fromPart = toLocalTimePart(fromDate);
  const toPart = toLocalTimePart(toDate);
  const dateStr = toLocalDateStr(toDate);
  return `${fromPart} - ${toPart} ${dateStr}.json`;
}

function fetchWithTimeout(url, options = {}, timeoutMs = FETCH_TIMEOUT_MS) {
  const controller = new AbortController();
  const to = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(url, { ...options, signal: controller.signal }).finally(() => clearTimeout(to));
}

/** Сервер вернул HTML (капча, прокси, страница входа) вместо JSON API */
function throwHtmlResponseError() {
  const e = new Error(
    'Сервер вернул страницу (HTML) вместо данных API. Возможные причины: капча/проверка на сайте, перехват корпоративным прокси или фаерволом. Откройте wwh.samokat.ru в браузере, при необходимости пройдите проверку; попробуйте другую сеть или режим (Офис/VPN).'
  );
  e.status = 502;
  throw e;
}

async function fetchOnePage(tokenToUse, bodyOptions, headers) {
  const body = buildBody(bodyOptions);
  let response;
  try {
    response = await fetchWithTimeout(API_URL, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
    });
  } catch (err) {
    if (err.name === 'AbortError') {
      const e = new Error(`Таймаут запроса к API (${FETCH_TIMEOUT_MS / 1000} с). Проверьте сеть или прокси.`);
      e.status = 408;
      throw e;
    }
    const code = err.cause && err.cause.code;
    const msg = code === 'ENOTFOUND' ? 'Сервер не найден (DNS). Проверьте интернет или VPN.' : (err.message || String(err));
    const e = new Error(`Сеть: ${msg}`);
    e.status = 502;
    throw e;
  }
  const text = await response.text();
  const trimmed = text && typeof text === 'string' ? text.trim().toLowerCase().replace(/\s+/g, ' ') : '';
  const isHtml = trimmed.startsWith('<!doctype') || trimmed.startsWith('<html');
  let data;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    if (isHtml) throwHtmlResponseError();
    const e = new Error(`Ответ не JSON: ${text.slice(0, 200)}`);
    e.status = 502;
    throw e;
  }
  if (!response.ok) {
    const errMsg = data?.message || data?.error || response.statusText || text.slice(0, 200);
    const err = new Error(`API ${response.status}: ${errMsg}`);
    err.status = response.status;
    err.data = data;
    throw err;
  }
  const value = data?.value || data;
  const items = Array.isArray(value?.items) ? value.items : (Array.isArray(data?.content) ? data.content : []);
  const total = value?.total ?? data?.totalElements ?? null;
  return { data, items, total };
}

async function fetchFromAPI(token, bodyOptions = {}) {
  const config = loadConfig();
  const tokenToUse = token || config.token;
  if (!tokenToUse || !tokenToUse.trim()) {
    const e = new Error('Токен не задан. Укажите токен в настройках.');
    e.status = 401;
    throw e;
  }

  const headers = {
    ...DEFAULT_HEADERS,
    'Authorization': `Bearer ${tokenToUse.trim()}`,
  };
  const cookie = (config.cookie || '').trim();
  if (cookie) headers['Cookie'] = cookie;

  const fetchAllPages = !!bodyOptions.fetchAllPages;
  const maxRows = bodyOptions.maxRows != null ? Math.max(1, parseInt(bodyOptions.maxRows, 10) || 0) : null;
  const requestedPageSize = parseInt(bodyOptions.pageSize, 10) || 100;
  const pageSize = Math.min(500, Math.max(100, requestedPageSize));
  const CONCURRENCY = 10;

  const startedAt = Date.now();

  if (fetchAllPages || (maxRows != null && maxRows > 0)) {
    const baseOptions = { ...bodyOptions, pageSize };
    let totalFromApi = null;
    const pageResults = [];

    const fetchPage = async (pageNum) => {
      const { data, items, total } = await fetchOnePage(tokenToUse, { ...baseOptions, pageNumber: pageNum }, headers);
      return { pageNumber: pageNum, items, total };
    };

    const firstBatch = await Promise.all(
      Array.from({ length: CONCURRENCY }, (_, i) => fetchPage(i + 1))
    );
    for (const r of firstBatch) {
      if (r.total != null) totalFromApi = r.total;
      pageResults.push(r);
    }
    const totalPages = totalFromApi != null ? Math.ceil(totalFromApi / pageSize) : null;

    if (totalPages != null && totalPages > CONCURRENCY) {
      for (let start = CONCURRENCY + 1; start <= totalPages; start += CONCURRENCY) {
        const end = Math.min(start + CONCURRENCY - 1, totalPages);
        if (maxRows != null && pageResults.reduce((s, r) => s + r.items.length, 0) >= maxRows) break;
        const batch = await Promise.all(
          Array.from({ length: end - start + 1 }, (_, i) => fetchPage(start + i))
        );
        for (const r of batch) {
          pageResults.push(r);
          if (r.items.length < pageSize) break;
        }
      }
    }

    pageResults.sort((a, b) => a.pageNumber - b.pageNumber);
    const allItems = pageResults.flatMap(r => r.items);
    const resultItems = maxRows != null ? allItems.slice(0, maxRows) : allItems;
    const count = resultItems.length;
    const total = totalFromApi ?? count;

    ensureDataDir();
    let name = filenameFromItems(resultItems) || (() => {
      const body = buildBody(baseOptions);
      return filenameFromRange(body.operationCompletedAtFrom, body.operationCompletedAtTo);
    })();
    let filename = path.join(DATA_DIR, name);
    let n = 1;
    while (fs.existsSync(filename)) {
      name = name.replace(/\.json$/, `_${n}.json`);
      filename = path.join(DATA_DIR, name);
      n++;
    }
    const payload = { fetchedAt: new Date().toISOString(), value: { items: resultItems, total } };
    fs.writeFileSync(filename, JSON.stringify(payload, null, 2), 'utf8');

    const lightName = name.replace(/\.json$/, '.light.json');
    try {
      const summary = buildDashboardSummary(resultItems, path.join(__dirname, '..', 'empl.csv'));
      summary.meta.fetchedAt = new Date().toISOString();
      summary.meta.sourceFile = name;
      fs.writeFileSync(path.join(DATA_DIR, lightName), JSON.stringify(summary, null, 2), 'utf8');
    } catch (err) {
      console.error('Не удалось записать лёгкий файл', lightName, err.message);
    }

    return {
      success: true,
      data: { value: { items: resultItems, total }, raw: payload },
      count,
      total,
      duration: Date.now() - startedAt,
      status: 200,
      savedTo: filename,
      pagesFetched: pageResults.length,
    };
  }

  const body = buildBody({ ...bodyOptions, pageNumber: bodyOptions.pageNumber || 1, pageSize });
  let response;
  try {
    response = await fetchWithTimeout(API_URL, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
    });
  } catch (err) {
    if (err.name === 'AbortError') {
      const e = new Error(`Таймаут запроса к API (${FETCH_TIMEOUT_MS / 1000} с). Проверьте сеть или прокси.`);
      e.status = 408;
      throw e;
    }
    const code = err.cause && err.cause.code;
    const msg = code === 'ENOTFOUND' ? 'Сервер не найден (DNS). Проверьте интернет или VPN.' : (err.message || String(err));
    const e = new Error(`Сеть: ${msg}`);
    e.status = 502;
    throw e;
  }

  const duration = Date.now() - startedAt;
  const text = await response.text();
  const trimmed = text && typeof text === 'string' ? text.trim().toLowerCase().replace(/\s+/g, ' ') : '';
  const isHtml = trimmed.startsWith('<!doctype') || trimmed.startsWith('<html');
  let data;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    if (isHtml) throwHtmlResponseError();
    const e = new Error(`Ответ не JSON: ${text.slice(0, 200)}`);
    e.status = 502;
    throw e;
  }

  if (!response.ok) {
    const errMsg = data?.message || data?.error || response.statusText || text.slice(0, 200);
    const err = new Error(`API ${response.status}: ${errMsg}`);
    err.status = response.status;
    err.data = data;
    throw err;
  }

  const value = data?.value || data;
  const items = Array.isArray(value?.items) ? value.items : (Array.isArray(data?.content) ? data.content : []);
  const count = items.length;
  const total = value?.total ?? data?.totalElements ?? count;

  ensureDataDir();
  let name = filenameFromItems(items) || filenameFromRange(body.operationCompletedAtFrom, body.operationCompletedAtTo);
  let filename = path.join(DATA_DIR, name);
  let n = 1;
  while (fs.existsSync(filename)) {
    name = name.replace(/\.json$/, `_${n}.json`);
    filename = path.join(DATA_DIR, name);
    n++;
  }
  fs.writeFileSync(filename, JSON.stringify({ fetchedAt: new Date().toISOString(), ...data }, null, 2), 'utf8');

  const lightName = name.replace(/\.json$/, '.light.json');
  const lightPath = path.join(DATA_DIR, lightName);
  try {
    const summary = buildDashboardSummary(items, path.join(__dirname, '..', 'empl.csv'));
    summary.meta.fetchedAt = new Date().toISOString();
    summary.meta.sourceFile = name;
    fs.writeFileSync(lightPath, JSON.stringify(summary, null, 2), 'utf8');
  } catch (err) {
    console.error('Не удалось записать лёгкий файл', lightName, err.message);
  }

  return {
    success: true,
    data: { value: { items, total }, raw: data },
    count,
    total,
    duration,
    status: response.status,
    savedTo: filename,
  };
}

module.exports = {
  fetchFromAPI,
  buildBody,
  filenameFromRange,
  filenameFromItems,
  API_URL,
};
