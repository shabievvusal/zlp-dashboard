'use strict';

const fs = require('fs');
const path = require('path');

const IDLE_THRESHOLD_MS = 10 * 60 * 1000; // 10 минут
const DSH_HOURS = Array.from({ length: 12 }, (_, i) => 10 + i); // 10..21 (день: 9:00–10:00 … 20:00–21:00)
const DSH_HOURS_NIGHT = [22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9]; // ночь: 22:00–23:00 … 8:00–9:00

function isDayShift(row) {
  if (!row || !row.operationCompletedAt) return false;
  const h = new Date(row.operationCompletedAt).getHours();
  return h >= 9 && h < 21;
}

function isNightShift(row) {
  if (!row || !row.operationCompletedAt) return false;
  const h = new Date(row.operationCompletedAt).getHours();
  return h >= 21 || h < 9;
}

function flattenItem(item) {
  if (!item || typeof item !== 'object') return item;
  const p = item.product || {};
  const ru = item.responsibleUser || {};
  const src = item.sourceAddress || {};
  const tgt = item.targetAddress || {};
  const part = item.part || {};
  return {
    id: item.id || '—',
    type: item.type || '—',
    operationType: item.operationType || '—',
    productName: p.name || '—',
    nomenclatureCode: p.nomenclatureCode || '—',
    productId: p.productId || item.productId || '—',
    targetCellAddress: (tgt && tgt.cellAddress) || item.targetCellAddress || '—',
    operationCompletedAt: item.operationCompletedAt || null,
    responsibleUser: [ru.lastName, ru.firstName, ru.middleName].filter(Boolean).join(' ') || '—',
    targetNew: (item.targetQuantity && item.targetQuantity.newQuantity) != null ? item.targetQuantity.newQuantity : 0,
    targetBarcode: (tgt && tgt.handlingUnitBarcode) || '—',
  };
}

function formatTimeOnly(isoStr) {
  if (!isoStr) return '—';
  const d = new Date(isoStr);
  const h = String(d.getHours()).padStart(2, '0');
  const m = String(d.getMinutes()).padStart(2, '0');
  return `${h}:${m}`;
}

function normalizeFio(s) {
  return (s || '').trim().replace(/\s+/g, ' ');
}

function fioToKey(fio) {
  const parts = normalizeFio(fio).split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).join(' ');
}

function buildEmplMap(emplCsvPath) {
  const map = new Map();
  if (!emplCsvPath || !fs.existsSync(emplCsvPath)) return map;
  let text;
  try {
    const buf = fs.readFileSync(emplCsvPath);
    const iconv = require('iconv-lite');
    text = iconv.decode(buf, 'cp1251');
  } catch {
    try { text = fs.readFileSync(emplCsvPath, 'utf8'); } catch { return map; }
  }
  const lines = (text || '').replace(/\r\n/g, '\n').split('\n');
  for (const line of lines) {
    const t = line.trim();
    if (!t) continue;
    const idx = t.indexOf(';');
    if (idx < 0) continue;
    const fio = t.slice(0, idx).trim();
    const company = t.slice(idx + 1).trim();
    if (fio) map.set(fioToKey(fio), company || '—');
  }
  return map;
}

function getCompanyForFio(fio, emplMap) {
  if (!emplMap) return '—';
  const key = fioToKey(fio);
  return emplMap.get(key) ?? '—';
}

/** Строит массив stats по уже отфильтрованным flattened (например только день или только ночь). */
function buildStatsFromFlattened(flattened, emplMap) {
  if (!flattened || flattened.length === 0) return [];
  const byUser = new Map();
  for (const row of flattened) {
    let fio = (row.responsibleUser || '').trim();
    if (!fio || fio === '—' || fio === '') fio = 'Неизвестно';
    if (!byUser.has(fio)) byUser.set(fio, []);
    byUser.get(fio).push(row);
  }
  const stats = [];
  for (const [fio, rows] of byUser) {
    const company = getCompanyForFio(fio, emplMap);
    let pieceSelectionCount = 0;
    const uniqueKdkTasksSet = new Set();
    let pieces = 0;
    const eoSet = new Set();
    for (const r of rows) {
      if (r.operationType === 'PIECE_SELECTION_PICKING') pieceSelectionCount++;
      if (r.operationType === 'PICK_BY_LINE') {
        const productId = r.productId || 'no-product';
        const targetCell = r.targetCellAddress || 'no-target-cell';
        uniqueKdkTasksSet.add(`${productId}||${targetCell}`);
      }
      pieces += Number(r.targetNew) || 0;
      if (r.targetBarcode && r.targetBarcode.trim() && r.targetBarcode !== '—') {
        eoSet.add(r.targetBarcode.trim());
      }
    }
    const kdk = uniqueKdkTasksSet.size;
    const hr = pieceSelectionCount;
    const sz = pieceSelectionCount + kdk;
    const eo = eoSet.size;
    const sorted = [...rows].sort((a, b) => {
      const aT = a.operationCompletedAt ? new Date(a.operationCompletedAt).getTime() : 0;
      const bT = b.operationCompletedAt ? new Date(b.operationCompletedAt).getTime() : 0;
      return aT - bT;
    });
    const idles = [];
    for (let i = 1; i < sorted.length; i++) {
      const prevEnd = sorted[i - 1].operationCompletedAt ? new Date(sorted[i - 1].operationCompletedAt).getTime() : 0;
      const nextStart = sorted[i].operationCompletedAt ? new Date(sorted[i].operationCompletedAt).getTime() : 0;
      if (prevEnd && nextStart && nextStart - prevEnd >= IDLE_THRESHOLD_MS) {
        idles.push(formatTimeOnly(sorted[i - 1].operationCompletedAt) + '–' + formatTimeOnly(sorted[i].operationCompletedAt));
      }
    }
    const idlesStr = idles.length ? idles.join(', ') : '—';
    stats.push({ fio, company, sz, hr, kdk, pieces, idlesStr, eo });
  }
  stats.sort((a, b) => (b.sz - a.sz) || (b.pieces - a.pieces));
  return stats;
}

/**
 * Строит саммари для дашборда: статика (суммы по ФИО) для быстрой отрисовки.
 * @param {Array} items - сырые items из API (value.items)
 * @param {string} [emplCsvPath] - путь к empl.csv для подстановки компании
 * @returns {{ meta: object, stats: Array, statsDay: Array, statsNight: Array, szByHour: object, szByHourNight: object }}
 */
function buildDashboardSummary(items, emplCsvPath) {
  const emplMap = buildEmplMap(emplCsvPath);
  if (!Array.isArray(items) || items.length === 0) {
    return { meta: { recordCount: 0 }, stats: [], statsDay: [], statsNight: [], szByHour: { hours: DSH_HOURS, rows: [] }, szByHourNight: { hours: DSH_HOURS_NIGHT, rows: [] }, type: 'light' };
  }
  const flattened = items.map(flattenItem);
  const flattenedDay = flattened.filter(isDayShift);
  const flattenedNight = flattened.filter(isNightShift);

  const stats = buildStatsFromFlattened(flattened, emplMap);
  const statsDay = buildStatsFromFlattened(flattenedDay, emplMap);
  const statsNight = buildStatsFromFlattened(flattenedNight, emplMap);

  const szByHour = computeSzByHour(flattened, emplMap, DSH_HOURS, (hour) => hour >= 9 && hour < 21, getColDay);
  const szByHourNight = computeSzByHour(flattened, emplMap, DSH_HOURS_NIGHT, (hour) => hour >= 21 || hour < 9, getColNight);

  const firstAt = flattened.reduce((acc, r) => {
    const t = r.operationCompletedAt ? new Date(r.operationCompletedAt).getTime() : null;
    return t != null && (acc == null || t < acc) ? t : acc;
  }, null);
  const lastAt = flattened.reduce((acc, r) => {
    const t = r.operationCompletedAt ? new Date(r.operationCompletedAt).getTime() : null;
    return t != null && (acc == null || t > acc) ? t : acc;
  }, null);
  return {
    meta: {
      recordCount: flattened.length,
      firstAt: firstAt != null ? new Date(firstAt).toISOString() : null,
      lastAt: lastAt != null ? new Date(lastAt).toISOString() : null,
    },
    stats,
    statsDay,
    statsNight,
    szByHour,
    szByHourNight,
    type: 'light',
  };
}

/** getColDay: час 9→10, 20→21. getColNight: 21→22, 22→23, 23→0, 0→1, …, 8→9. */
function getColDay(hour) { return hour + 1; }
function getColNight(hour) { return hour === 23 ? 0 : hour + 1; }

function computeSzByHour(flattened, emplMap, hoursArray, hourFilter, getCol) {
  if (!flattened || flattened.length === 0) {
    return { hours: hoursArray, rows: [] };
  }
  if (!getCol) getCol = getColDay;
  const byFio = new Map();
  for (const row of flattened) {
    if (!row.operationCompletedAt) continue;
    const dt = new Date(row.operationCompletedAt);
    const hour = dt.getHours();
    if (!hourFilter(hour)) continue;
    const col = getCol(hour);
    let fio = (row.responsibleUser || '').trim();
    if (!fio || fio === '—') fio = 'Неизвестно';
    if (!byFio.has(fio)) byFio.set(fio, new Map());
    const byHour = byFio.get(fio);
    if (!byHour.has(col)) byHour.set(col, { pieceSelectionCount: 0, kdkSet: new Set() });
    const cell = byHour.get(col);
    if (row.operationType === 'PIECE_SELECTION_PICKING') {
      cell.pieceSelectionCount++;
    } else if (row.operationType === 'PICK_BY_LINE') {
      const productId = row.productId || 'no-product';
      const targetCell = row.targetCellAddress || 'no-target-cell';
      cell.kdkSet.add(`${productId}||${targetCell}`);
    }
  }
  const getTotal = (fio) => {
    const byHour = byFio.get(fio);
    if (!byHour) return 0;
    return hoursArray.reduce((s, h) => {
      const cell = byHour.get(h);
      return s + (cell?.pieceSelectionCount || 0) + (cell?.kdkSet?.size || 0);
    }, 0);
  };
  const fios = [...byFio.keys()].sort((a, b) => {
    const companyA = getCompanyForFio(a, emplMap) || '—';
    const companyB = getCompanyForFio(b, emplMap) || '—';
    const byCompany = companyA.localeCompare(companyB);
    if (byCompany !== 0) return byCompany;
    return getTotal(b) - getTotal(a);
  });
  const rows = fios.map((fio) => {
    const company = getCompanyForFio(fio, emplMap) || '—';
    const byHour = byFio.get(fio);
    const values = hoursArray.map((h) => {
      const cell = byHour?.get(h);
      return (cell?.pieceSelectionCount || 0) + (cell?.kdkSet?.size || 0);
    });
    const total = values.reduce((a, b) => a + b, 0);
    return { fio, company, values, total };
  });
  return { hours: hoursArray, rows };
}

module.exports = {
  buildDashboardSummary,
  buildEmplMap,
};
