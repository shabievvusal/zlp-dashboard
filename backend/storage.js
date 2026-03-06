/**
 * storage.js — хранилище операций для страницы /vs
 * Почасовые файлы: data/YYYY-MM-DD/HH.json (лёгкий формат, только нужные поля)
 * Обратная совместимость: при отсутствии почасовых данных читаем shift_YYYY-MM-DD_day|night.json
 */

const fs = require('fs');
const path = require('path');

const DATA_DIR = path.join(__dirname, 'data');

function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
}

function getShiftKey(isoDate) {
  const d = new Date(isoDate);
  const h = d.getHours();
  if (h >= 9 && h < 21) {
    const dateStr = d.toISOString().slice(0, 10);
    return `${dateStr}_day`;
  } else {
    const base = new Date(d);
    if (h < 9) base.setDate(base.getDate() - 1);
    const dateStr = base.toISOString().slice(0, 10);
    return `${dateStr}_night`;
  }
}

function getCurrentShiftKey() {
  return getShiftKey(new Date().toISOString());
}

/** Получить mergeKey из полного объекта операции (API) */
function getMergeKey(item) {
  const type = (item.operationType || item.type || '').toUpperCase();
  const isTaskType = type === 'PICK_BY_LINE' || type === 'PIECE_SELECTION_PICKING';
  if (isTaskType) {
    const exec = (item.responsibleUser && (item.responsibleUser.id || [item.responsibleUser.lastName, item.responsibleUser.firstName].filter(Boolean).join(' '))) || '';
    const cell = (item.targetAddress && item.targetAddress.cellAddress) || (item.sourceAddress && item.sourceAddress.cellAddress) || '';
    const product = (item.product && (item.product.nomenclatureCode || item.product.name)) || '';
    return `task|${exec}|${cell}|${product}`;
  }
  return `id|${item.id || ''}`;
}

/** MergeKey для уже облегчённого объекта (поля верхнего уровня) */
function getMergeKeyFromLight(light) {
  const type = (light.operationType || light.type || '').toUpperCase();
  const isTaskType = type === 'PICK_BY_LINE' || type === 'PIECE_SELECTION_PICKING';
  if (isTaskType) {
    const exec = light.executor || '';
    const cell = light.cell || '';
    const product = light.nomenclatureCode || light.productName || '';
    return `task|${exec}|${cell}|${product}`;
  }
  return `id|${light.id || ''}`;
}

/** Привести полный объект операции с API к лёгкому формату (как flattenItem на клиенте) */
function toLightItem(item) {
  const ru = item.responsibleUser || {};
  const executor = [ru.lastName, ru.firstName].filter(Boolean).join(' ').trim() || '';
  const product = item.product || {};
  return {
    id: item.id || '',
    type: item.type || '',
    operationType: item.operationType || '',
    productName: product.name || '',
    nomenclatureCode: product.nomenclatureCode || '',
    barcodes: (product.barcodes || []).join(', '),
    productionDate: item.part?.productionDate || '',
    bestBeforeDate: item.part?.bestBeforeDate || '',
    sourceBarcode: item.sourceAddress?.handlingUnitBarcode || '',
    cell: (item.targetAddress && item.targetAddress.cellAddress) || (item.sourceAddress && item.sourceAddress.cellAddress) || '',
    targetBarcode: item.targetAddress?.handlingUnitBarcode || '',
    startedAt: item.operationStartedAt || '',
    completedAt: item.operationCompletedAt || '',
    executor,
    executorId: ru.id || '',
    srcOld: item.sourceQuantity?.oldQuantity ?? '',
    srcNew: item.sourceQuantity?.newQuantity ?? '',
    tgtOld: item.targetQuantity?.oldQuantity ?? '',
    tgtNew: item.targetQuantity?.newQuantity ?? '',
    quantity: item.targetQuantity?.newQuantity ?? item.sourceQuantity?.oldQuantity ?? '',
  };
}

// ─── Почасовое хранение (лёгкий формат) ─────────────────────────────────────

function hourlyDir(dateStr) {
  return path.join(DATA_DIR, dateStr);
}

function hourlyFilePath(dateStr, hour) {
  return path.join(hourlyDir(dateStr), `${String(hour).padStart(2, '0')}.json`);
}

function ensureHourlyDir(dateStr) {
  const dir = hourlyDir(dateStr);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return dir;
}

/** Загрузить один почасовой файл. Возвращает Map(mergeKey -> lightItem). */
function loadHourly(dateStr, hour) {
  const fp = hourlyFilePath(dateStr, hour);
  if (!fs.existsSync(fp)) return new Map();
  try {
    const raw = JSON.parse(fs.readFileSync(fp, 'utf8'));
    const items = new Map();
    const list = Array.isArray(raw.items) ? raw.items : Object.values(raw.items || {});
    for (const item of list) {
      const k = getMergeKeyFromLight(item);
      if (!items.has(k)) items.set(k, item);
    }
    return items;
  } catch {
    return new Map();
  }
}

/** Сохранить один почасовой файл (только лёгкие объекты). */
function saveHourly(dateStr, hour, itemsMap) {
  ensureDataDir();
  ensureHourlyDir(dateStr);
  const fp = hourlyFilePath(dateStr, hour);
  const items = Array.from(itemsMap.values());
  const obj = {
    date: dateStr,
    hour: Number(hour),
    updatedAt: new Date().toISOString(),
    items,
  };
  fs.writeFileSync(fp, JSON.stringify(obj), 'utf8');
}

/** Есть ли почасовые данные за дату (хотя бы один файл). */
function hasHourlyDataForDate(dateStr) {
  const dir = hourlyDir(dateStr);
  if (!fs.existsSync(dir)) return false;
  const files = fs.readdirSync(dir).filter(f => /^\d{2}\.json$/.test(f));
  return files.length > 0;
}

/** Есть ли почасовые данные за предыдущую дату (для ночной смены). */
function hasAnyHourlyData(dateStr) {
  const prev = new Date(dateStr);
  prev.setDate(prev.getDate() - 1);
  const prevStr = prev.toISOString().slice(0, 10);
  return hasHourlyDataForDate(dateStr) || hasHourlyDataForDate(prevStr);
}

// ─── Мерж: сохраняем по часам в лёгком формате ──────────────────────────────

function mergeOperations(newItems) {
  if (!Array.isArray(newItems) || newItems.length === 0) {
    return { added: 0, skipped: 0, byShift: {} };
  }
  const byDateHour = new Map();
  for (const item of newItems) {
    const ts = item.operationCompletedAt;
    if (!ts) continue;
    const d = new Date(ts);
    const dateStr = d.toISOString().slice(0, 10);
    const hour = d.getHours();
    const key = `${dateStr}\t${hour}`;
    if (!byDateHour.has(key)) byDateHour.set(key, []);
    byDateHour.get(key).push(item);
  }

  let totalAdded = 0;
  let totalSkipped = 0;
  const byShift = {};

  for (const [dateHourKey, items] of byDateHour) {
    const [dateStr, hourStr] = dateHourKey.split('\t');
    const hour = parseInt(hourStr, 10);
    const shiftKey = getShiftKey(new Date(`${dateStr}T${String(hour).padStart(2, '0')}:00:00`));
    if (!byShift[shiftKey]) byShift[shiftKey] = { added: 0, skipped: 0, total: 0 };

    const existing = loadHourly(dateStr, hour);
    let added = 0;
    let skipped = 0;
    for (const item of items) {
      const light = toLightItem(item);
      const mergeKey = getMergeKey(item);
      if (existing.has(mergeKey)) skipped++;
      else {
        existing.set(mergeKey, light);
        added++;
      }
    }
    saveHourly(dateStr, hour, existing);
    byShift[shiftKey].added += added;
    byShift[shiftKey].skipped += skipped;
    byShift[shiftKey].total = existing.size;
    totalAdded += added;
    totalSkipped += skipped;
  }

  return { added: totalAdded, skipped: totalSkipped, byShift };
}

// ─── Чтение: почасовые файлы или fallback на смены ───────────────────────────

/** Список (dateStr, hour) для загрузки. Всегда отдаём все часы за дату (0–23 + ночь пред. дня), фильтр смены — на клиенте по локальному времени (UTC 06:00 = 09:00 МСК и т.д.). */
function getHoursToLoad(dateStr, fromHour, toHour, shift) {
  const prev = new Date(dateStr);
  prev.setDate(prev.getDate() - 1);
  const prevStr = prev.toISOString().slice(0, 10);
  const pairs = [];

  if (fromHour !== undefined || toHour !== undefined) {
    const from = fromHour == null ? 0 : Math.max(0, fromHour);
    const to = toHour == null ? 23 : Math.min(23, toHour);
    for (let h = from; h <= to; h++) pairs.push([dateStr, h]);
    return pairs;
  }
  // Полный день: ночь предыдущего (21–23) + все часы текущей даты (0–23)
  for (const h of [21, 22, 23]) pairs.push([prevStr, h]);
  for (let h = 0; h <= 23; h++) pairs.push([dateStr, h]);
  return pairs;
}

function getDateItemsFromHourly(dateStr, options = {}) {
  const { fromHour, toHour, shift } = options;
  const pairs = getHoursToLoad(dateStr, fromHour, toHour, shift);
  const byId = new Map();
  for (const [d, hour] of pairs) {
    const map = loadHourly(d, hour);
    for (const item of map.values()) {
      const k = item.id || (item.completedAt + item.executor + item.cell);
      if (!byId.has(k)) byId.set(k, item);
    }
  }
  const items = Array.from(byId.values());
  const ts = item => item.completedAt || item.startedAt || '';
  items.sort((a, b) => ts(a).localeCompare(ts(b)));
  return items;
}

function getDateItems(dateStr, options = {}) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return [];
  if (hasAnyHourlyData(dateStr)) {
    return getDateItemsFromHourly(dateStr, options);
  }
  const prevDate = new Date(dateStr);
  prevDate.setDate(prevDate.getDate() - 1);
  const prevDateStr = prevDate.toISOString().slice(0, 10);
  const nightShift = loadShift(prevDateStr + '_night');
  const dayShift = loadShift(dateStr + '_day');
  const byId = new Map();
  if (options.shift !== 'day') {
    for (const item of nightShift.items.values()) byId.set(item.id, item);
  }
  if (options.shift !== 'night') {
    for (const item of dayShift.items.values()) byId.set(item.id, item);
  }
  let items = Array.from(byId.values());
  items = items.map(i => (i.executor !== undefined ? i : toLightItem(i)));
  const ts = item => item.operationCompletedAt || item.operationStartedAt || '';
  items.sort((a, b) => ts(a).localeCompare(ts(b)));
  return items;
}

// ─── Старые смены (для совместимости) ────────────────────────────────────────

function shiftFilePath(shiftKey) {
  return path.join(DATA_DIR, `shift_${shiftKey}.json`);
}

function loadShift(shiftKey) {
  const fp = shiftFilePath(shiftKey);
  if (!fs.existsSync(fp)) {
    return { shiftKey, updatedAt: null, items: new Map() };
  }
  try {
    const raw = JSON.parse(fs.readFileSync(fp, 'utf8'));
    const items = new Map();
    for (const item of Object.values(raw.items || {})) {
      const k = getMergeKey(item);
      if (!items.has(k)) items.set(k, item);
    }
    return { shiftKey, updatedAt: raw.updatedAt || null, items };
  } catch {
    return { shiftKey, updatedAt: null, items: new Map() };
  }
}

function saveShift(shiftData) {
  ensureDataDir();
  const { shiftKey, items } = shiftData;
  const fp = shiftFilePath(shiftKey);
  const obj = {
    shiftKey,
    updatedAt: new Date().toISOString(),
    items: Object.fromEntries(items),
  };
  fs.writeFileSync(fp, JSON.stringify(obj), 'utf8');
}

function getShiftItems(shiftKey) {
  const shift = loadShift(shiftKey);
  return Array.from(shift.items.values());
}

function listShifts() {
  ensureDataDir();
  if (!fs.existsSync(DATA_DIR)) return [];
  const result = [];
  const dirs = fs.readdirSync(DATA_DIR, { withFileTypes: true }).filter(e => e.isDirectory() && /^\d{4}-\d{2}-\d{2}$/.test(e.name));
  for (const dir of dirs) {
    const dateStr = dir.name;
    const hourFiles = fs.readdirSync(path.join(DATA_DIR, dateStr)).filter(f => /^\d{2}\.json$/.test(f));
    let dayCount = 0;
    let nightCount = 0;
    let lastUpdated = null;
    for (const f of hourFiles) {
      const hour = parseInt(f.replace('.json', ''), 10);
      const fp = path.join(DATA_DIR, dateStr, f);
      const stat = fs.statSync(fp);
      if (lastUpdated === null || stat.mtime > lastUpdated) lastUpdated = stat.mtime;
      try {
        const raw = JSON.parse(fs.readFileSync(fp, 'utf8'));
        const n = Array.isArray(raw.items) ? raw.items.length : Object.keys(raw.items || {}).length;
        if (hour >= 9 && hour < 21) dayCount += n;
        else if (hour >= 0 && hour < 9) nightCount += n;
      } catch {}
    }
    const prev = new Date(dateStr);
    prev.setDate(prev.getDate() - 1);
    const prevStr = prev.toISOString().slice(0, 10);
    if (fs.existsSync(hourlyDir(prevStr))) {
      for (const h of [21, 22, 23]) {
        const m = loadHourly(prevStr, h);
        nightCount += m.size;
      }
    }
    if (dayCount > 0) result.push({ shiftKey: `${dateStr}_day`, date: dateStr, type: 'day', count: dayCount, updatedAt: lastUpdated ? new Date(lastUpdated).toISOString() : null, fileSize: null });
    if (nightCount > 0) result.push({ shiftKey: `${dateStr}_night`, date: dateStr, type: 'night', count: nightCount, updatedAt: lastUpdated ? new Date(lastUpdated).toISOString() : null, fileSize: null });
  }
  const shiftFiles = fs.readdirSync(DATA_DIR).filter(f => f.startsWith('shift_') && f.endsWith('.json'));
  for (const f of shiftFiles) {
    const shiftKey = f.replace('shift_', '').replace('.json', '');
    if (result.some(r => r.shiftKey === shiftKey)) continue;
    const stat = fs.statSync(path.join(DATA_DIR, f));
    const shift = loadShift(shiftKey);
    result.push({
      shiftKey,
      date: shiftKey.split('_')[0],
      type: shiftKey.split('_')[1],
      count: shift.items.size,
      updatedAt: shift.updatedAt,
      fileSize: stat.size,
    });
  }
  return result.sort((a, b) => (b.shiftKey || '').localeCompare(a.shiftKey || ''));
}

module.exports = {
  mergeOperations,
  getShiftItems,
  getDateItems,
  listShifts,
  getCurrentShiftKey,
  getShiftKey,
  DATA_DIR,
  ensureDataDir,
  toLightItem,
  getMergeKeyFromLight,
};
