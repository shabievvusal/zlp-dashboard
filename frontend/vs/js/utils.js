/**
 * utils.js — вспомогательные функции
 */

export const el = id => document.getElementById(id);

export function formatDateTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export function formatDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`;
}

export function formatTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getHours())}:${p(d.getMinutes())}`;
}

/**
 * Разворачивает вложенный объект операции в плоский.
 */
export function flattenItem(item) {
  return {
    id: item.id || '',
    type: item.type || '',
    operationType: item.operationType || '',
    productName: item.product?.name || '',
    nomenclatureCode: item.product?.nomenclatureCode || '',
    barcodes: (item.product?.barcodes || []).join(', '),
    productionDate: item.part?.productionDate || '',
    bestBeforeDate: item.part?.bestBeforeDate || '',
    sourceBarcode: item.sourceAddress?.handlingUnitBarcode || '',
    cell: item.targetAddress?.cellAddress || item.sourceAddress?.cellAddress || '',
    targetBarcode: item.targetAddress?.handlingUnitBarcode || '',
    startedAt: item.operationStartedAt || '',
    completedAt: item.operationCompletedAt || '',
    executor: item.responsibleUser
      ? `${item.responsibleUser.lastName || ''} ${item.responsibleUser.firstName || ''}`.trim()
      : '',
    executorId: item.responsibleUser?.id || '',
    srcOld: item.sourceQuantity?.oldQuantity ?? '',
    srcNew: item.sourceQuantity?.newQuantity ?? '',
    tgtOld: item.targetQuantity?.oldQuantity ?? '',
    tgtNew: item.targetQuantity?.newQuantity ?? '',
    quantity: item.targetQuantity?.newQuantity ?? item.sourceQuantity?.oldQuantity ?? '',
  };
}

/**
 * Парсит CSV сотрудников → Map(normalizedFio → company)
 */
export function parseEmplCsv(csv) {
  const map = new Map();
  const companies = new Set();
  if (!csv) return { map, companies: [] };

  const lines = csv.split('\n');
  // Определяем разделитель по первой строке
  const firstLine = lines[0] || '';
  const sep = firstLine.includes(';') ? ';' : ',';

  // Определяем, есть ли строка-заголовок:
  // считаем заголовком, если первая колонка — «фио», «name», «имя» и т.п.
  const firstCol = firstLine.split(sep)[0].trim().toLowerCase().replace(/^"|"$/g, '');
  const looksLikeHeader = ['фио', 'имя', 'name', 'сотрудник', 'ф.и.о', 'fio'].includes(firstCol);
  const startRow = looksLikeHeader ? 1 : 0;

  for (let i = startRow; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const cols = line.split(sep).map(c => c.trim().replace(/^"|"$/g, ''));
    const fio = cols[0] || '';
    const company = cols[1] || '';
    if (fio) {
      map.set(normalizeFio(fio), company);
      if (company) companies.add(company);
    }
  }
  return { map, companies: [...companies].sort() };
}

export function normalizeFio(fio) {
  return (fio || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Ключ для сопоставления короткого ФИО (Иванов И.) с полным: фамилия + первая буква имени */
export function personKey(norm) {
  const parts = (norm || '').split(/\s+/).filter(Boolean);
  if (parts.length === 0) return norm;
  const initial = parts.length > 1 ? parts[1].charAt(0).toLowerCase() : '';
  return (parts[0] + ' ' + initial).trim();
}

/** Есть ли в emplMap запись по точному ключу или по personKey */
export function hasMatchInEmplKeys(dataNorm, emplMap) {
  if (!dataNorm || !emplMap) return false;
  if (emplMap.has(dataNorm)) return true;
  const pk = personKey(dataNorm);
  for (const k of emplMap.keys()) {
    if (personKey(k) === pk) return true;
  }
  return false;
}

/** Компания по ФИО из данных: точный ключ в emplMap или по personKey */
export function getCompanyByFio(emplMap, dataNorm) {
  if (!emplMap || !dataNorm) return undefined;
  const exact = emplMap.get(dataNorm);
  if (exact !== undefined) return exact;
  const pk = personKey(dataNorm);
  for (const [k, v] of emplMap) {
    if (personKey(k) === pk) return v;
  }
  return undefined;
}

/**
 * Экспорт массива объектов в CSV.
 */
export function exportToCsv(rows, filename) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const lines = [
    headers.join(';'),
    ...rows.map(r => headers.map(h => `"${String(r[h] ?? '').replace(/"/g, '""')}"`).join(';')),
  ];
  const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function shiftLabel(shiftKey) {
  if (!shiftKey) return '';
  const [date, type] = shiftKey.split('_');
  const [y, m, d] = (date || '').split('-');
  const dateStr = d && m && y ? `${d}.${m}.${y}` : date;
  return type === 'day' ? `${dateStr} День (9:00–21:00)` : `${dateStr} Ночь (21:00–9:00)`;
}
