/**
 * stats.js — подсчёт и отображение статистики по операциям
 */

import { el, normalizeFio, formatTime, hasMatchInEmplKeys, getCompanyByFio } from './utils.js';

/**
 * Ключ "задачи": для КДК (По линии) — один вклад в одну ячейку одним товаром = одна задача; для остальных — одна операция = одна задача.
 */
function getTaskKey(item) {
  const type = (item.operationType || '').toUpperCase();
  if (type === 'PICK_BY_LINE') {
    const exec = item.executorId || item.executor || '';
    const cell = item.cell || '';
    const product = item.nomenclatureCode || item.productName || '';
    return `kdk|${exec}|${cell}|${product}`;
  }
  return item.id ? `op|${item.id}` : `op|${(item.completedAt || item.startedAt || '')}|${item.executor || ''}|${item.cell || ''}`;
}

/**
 * Считает статистику по плоскому массиву операций.
 * Для КДК (По линии) несколько вкладов одного товара в одну ячейку одним сотрудником считаются одной задачей.
 * @param {Array} items — flattenItem[]
 * @param {Map} emplMap — Map(normalizedFio -> company)
 * @param {string} filterCompany — '__all__' | '__none__' | company
 */
export function calcStats(items, emplMap, filterCompany) {
  const filtered = filterByCompany(items, emplMap, filterCompany);

  const totalTaskKeys = new Set(filtered.map(i => getTaskKey(i)));
  const totalOps = totalTaskKeys.size;
  const totalQty = filtered.reduce((s, i) => s + (Number(i.quantity) || 0), 0);

  // Статистика по сотрудникам (ops = число задач с дедупом КДК)
  const byExecutor = new Map();
  for (const item of filtered) {
    const key = item.executor || 'Неизвестно';
    if (!byExecutor.has(key)) byExecutor.set(key, { name: key, taskKeys: new Set(), qty: 0, firstAt: null, lastAt: null });
    const e = byExecutor.get(key);
    e.taskKeys.add(getTaskKey(item));
    e.qty += Number(item.quantity) || 0;
    const ts = item.completedAt || item.startedAt;
    if (ts) {
      if (!e.firstAt || ts < e.firstAt) e.firstAt = ts;
      if (!e.lastAt  || ts > e.lastAt)  e.lastAt  = ts;
    }
  }
  const executors = [...byExecutor.values()].map(e => ({
    ...e,
    ops: e.taskKeys.size,
    company: emplMap ? (getCompanyByFio(emplMap, normalizeFio(e.name)) || '—') : '—',
  })).sort((a, b) => b.ops - a.ops);

  // Статистика по часам: ориентир — completedAt (время подтверждения задачи). Как на бэкенде.
  const byHour = new Map(); // hour -> { hour, taskKeys: Set, kdkTaskKeys: Set, employees: Set, storageOps, kdkOps }
  for (const item of filtered) {
    const ts = item.completedAt;
    if (!ts) continue;
    const h = new Date(ts).getHours();
    if (!byHour.has(h)) byHour.set(h, { hour: h, taskKeys: new Set(), kdkTaskKeys: new Set(), employees: new Set(), storageOps: 0, kdkOps: 0 });
    const hh = byHour.get(h);
    const type = (item.operationType || '').toUpperCase();
    const isKdk = type === 'PICK_BY_LINE';
    const tk = getTaskKey(item);
    hh.taskKeys.add(tk);
    if (isKdk) hh.kdkTaskKeys.add(tk); else hh.storageOps++;
    hh.kdkOps = hh.kdkTaskKeys.size;
    if (item.executorId || item.executor) hh.employees.add(item.executorId || item.executor);
  }
  const hourly = [...byHour.values()].map(x => ({
    hour: x.hour,
    ops: x.taskKeys.size,
    employees: x.employees.size,
    storageOps: x.storageOps,
    kdkOps: x.kdkOps,
  })).sort((a, b) => a.hour - b.hour);

  // Время старта и последнего пика (по completedAt)
  let firstAt = null;
  let lastAt = null;
  for (const item of filtered) {
    const ts = item.completedAt;
    if (!ts) continue;
    if (!firstAt || ts < firstAt) firstAt = ts;
    if (!lastAt  || ts > lastAt)  lastAt  = ts;
  }

  return { totalOps, totalQty, executors, filteredCount: filtered.length, hourly, firstAt, lastAt };
}

function filterByCompany(items, emplMap, filterCompany) {
  if (!emplMap || !filterCompany || filterCompany === '__all__') return items;
  if (filterCompany === '__none__') {
    return items.filter(i => !hasMatchInEmplKeys(normalizeFio(i.executor), emplMap));
  }
  return items.filter(i => getCompanyByFio(emplMap, normalizeFio(i.executor)) === filterCompany);
}

/**
 * Рендерит карточки статистики.
 */
export function renderStats(stats, shiftLabel) {
  const container = el('stats-cards');
  if (!container) return;

  container.innerHTML = `
    <div class="stat-card">
      <div class="stat-icon">📦</div>
      <div class="stat-value">${stats.totalOps.toLocaleString('ru-RU')}</div>
      <div class="stat-label">Операций</div>
    </div>
    <div class="stat-card stat-card--green">
      <div class="stat-icon">🔢</div>
      <div class="stat-value">${stats.totalQty.toLocaleString('ru-RU')}</div>
      <div class="stat-label">Единиц товара</div>
    </div>
    <div class="stat-card">
      <div class="stat-icon">👷</div>
      <div class="stat-value">${stats.executors.length}</div>
      <div class="stat-label">Сотрудников</div>
    </div>
    <div class="stat-card">
      <div class="stat-icon">📅</div>
      <div class="stat-value stat-value--sm">${shiftLabel || '—'}</div>
      <div class="stat-label">Дата</div>
    </div>

  `;
}

/**
 * Рендерит таблицу топ-сотрудников.
 */
export function renderExecutorTable(executors) {
  const tbody = el('executor-tbody');
  if (!tbody) return;

  if (!executors.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="empty-row">Нет данных</td></tr>';
    return;
  }

  const maxOps = Math.max(...executors.map(e => e.ops), 1);
  tbody.innerHTML = executors.map((e, i) => `
    <tr>
      <td class="rank">${i + 1}</td>
      <td class="executor-company">${escHtml(e.company || '—')}</td>
      <td class="executor-name">${escHtml(e.name)}</td>
      <td class="text-right">${e.qty.toLocaleString('ru-RU')}</td>
      <td class="qty-cell">
        <div class="qty-bar-wrap">
          <div class="qty-bar" style="width:${Math.round((e.ops / maxOps) * 100)}%"></div>
          <span class="qty-value">${e.ops.toLocaleString('ru-RU')}</span>
        </div>
      </td>
      <td class="text-right time-cell">${e.firstAt ? formatTime(e.firstAt) : '—'} – ${e.lastAt ? formatTime(e.lastAt) : '—'}</td>
    </tr>
  `).join('');
}

/** Часы для отображения: день — колонка 10 = 09:00–10:00, колонка 21 = 20:00–21:00 (номер колонки = конец часа) */
const DAY_HOURS = [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];
/** Ночь: колонка 23 = 22:00–23:00, 0 = 23:00–00:00, … 10 = 09:00–10:00 */
const NIGHT_HOURS = [23, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

/**
 * Приводит массив по часам к порядку и диапазону смены; заполняет нулями отсутствующие часы.
 * Метки столбцов = конец интервала (10 = 09:00–10:00, 21 = 20:00–21:00). Данные из calcStats по началу часа (9..20).
 * @param {Array} hourly — массив { hour, ops, employees, storageOps, kdkOps }
 * @param {'day'|'night'} shiftFilter
 */
export function getHourlyForShift(hourly, shiftFilter) {
  const byHour = new Map();
  if (Array.isArray(hourly)) {
    for (const h of hourly) byHour.set(h.hour, {
      hour: h.hour,
      ops: h.ops || 0,
      employees: h.employees ?? 0,
      storageOps: h.storageOps ?? 0,
      kdkOps: h.kdkOps ?? 0,
    });
  }
  const order = shiftFilter === 'night' ? NIGHT_HOURS : DAY_HOURS;
  return order.map(col => {
    const dataHour = shiftFilter === 'day' ? col - 1 : (col - 1 + 24) % 24;
    const h = byHour.get(dataHour) || { hour: dataHour, ops: 0, employees: 0, storageOps: 0, kdkOps: 0 };
    return { ...h, hour: col };
  });
}

/**
 * Рендерит диаграмму пиков по часам: сверху операции и сотрудников, два столбика — хранение и КДК, значение внутри столбика.
 */
export function renderHourlyChart(hourly, shiftFilter = 'day') {
  const container = el('hourly-chart');
  if (!container) return;

  const ordered = getHourlyForShift(hourly || [], shiftFilter);
  const hasData = ordered.some(h => h.ops > 0 || h.storageOps > 0 || h.kdkOps > 0);

  if (!hasData) {
    container.innerHTML = '<div class="empty-row" style="padding:20px;text-align:center;color:var(--text-muted)">Нет данных</div>';
    return;
  }

  const maxBar = Math.max(...ordered.map(h => Math.max(h.storageOps, h.kdkOps)), 1);

  container.innerHTML = `
    <div class="hourly-bars">
      ${ordered.map(h => `
        <div class="hourly-col">
          <div class="hourly-values">
            <span class="hourly-ops">${h.ops} оп.</span>
            <span class="hourly-employees">${h.employees} чел.</span>
          </div>
          <div class="hourly-bar-wrap">
            <div class="hourly-bar-storage" style="height:${Math.round((h.storageOps / maxBar) * 100)}%" title="Хранение">
              <span class="hourly-bar-value">${h.storageOps}</span>
            </div>
            <div class="hourly-bar-kdk" style="height:${Math.round((h.kdkOps / maxBar) * 100)}%" title="КДК">
              <span class="hourly-bar-value">${h.kdkOps}</span>
            </div>
          </div>
          <div class="hourly-label">${String(h.hour).padStart(2, '0')}:00</div>
        </div>
      `).join('')}
    </div>
  `;
}

/**
 * Считает для каждого сотрудника СЗ по каждому часу — так же, как на dsh:
 * ХР = только PIECE_SELECTION_PICKING (каждая операция), КДК = уникальные по (товар + ячейка) для PICK_BY_LINE.
 * СЗ = ХР + КДК (без двойного учёта).
 * @param {Array} items — flattenItem[] уже отфильтрованные по смене/компании
 * @param {'day'|'night'} shiftFilter
 * @returns {{ hours: number[], rows: Array<{name:string, byHour:Object, total:number}> }}
 */
export function calcHourlyByEmployee(items, shiftFilter = 'day') {
  const order = shiftFilter === 'night' ? NIGHT_HOURS : DAY_HOURS;

  // byEmployee: name -> Map<hour, { pieceSelectionCount: number, kdkSet: Set<product||cell> }>
  const byEmployee = new Map();

  for (const item of items) {
    const ts = item.completedAt;
    if (!ts) continue;
    const h = new Date(ts).getHours();
    // Колонка 10 = 09:00–10:00 (час 9), колонка 21 = 20:00–21:00 (час 20) → ключ col = конец интервала
    const col = (h + 1) % 24;
    const name = item.executor || 'Неизвестно';

    if (!byEmployee.has(name)) byEmployee.set(name, new Map());
    const hourMap = byEmployee.get(name);

    if (!hourMap.has(col)) hourMap.set(col, { pieceSelectionCount: 0, kdkSet: new Set() });
    const cell = hourMap.get(col);

    const type = (item.operationType || '').toUpperCase();
    if (type === 'PIECE_SELECTION_PICKING') {
      cell.pieceSelectionCount++;
    } else if (type === 'PICK_BY_LINE') {
      const productId = item.nomenclatureCode || item.productName || 'no-product';
      const targetCell = item.cell || 'no-target-cell';
      cell.kdkSet.add(`${productId}||${targetCell}`);
    }
  }

  const rows = [];
  for (const [name, hourMap] of byEmployee) {
    const byHour = {};
    let total = 0;
    for (const col of order) {
      const cell = hourMap.get(col);
      if (!cell) { byHour[col] = 0; continue; }
      const sz = cell.pieceSelectionCount + (cell.kdkSet ? cell.kdkSet.size : 0);
      byHour[col] = sz;
      total += sz;
    }
    rows.push({ name, byHour, total });
  }

  return { hours: order, rows };
}

/**
 * Часы, которые уже наступили (для выбранной даты). Для «сегодня» — только прошедшие; для прошлой даты — все.
 */
export function filterHoursToPassed(selectedDate, shiftFilter) {
  const order = shiftFilter === 'night' ? NIGHT_HOURS : DAY_HOURS;
  const today = typeof selectedDate === 'string' && selectedDate === getTodayStr();
  if (!today) return order;
  const now = new Date();
  const currentHour = now.getHours();
  if (shiftFilter === 'day') {
    return order.filter(col => col <= currentHour);
  }
  return order.filter(col => col >= 22 || col <= currentHour);
}

function getTodayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * Данные «Сотрудники по часам»: только прошедшие часы, с компанией, сгруппированы по компании (для отправки в Telegram).
 */
export function getHourlyByEmployeeGroupedByCompany(items, shiftFilter, emplMap, selectedDate) {
  const { hours: allHours, rows } = calcHourlyByEmployee(items, shiftFilter);
  const hours = filterHoursToPassed(selectedDate, shiftFilter);
  const getCompany = (name) => (emplMap && name ? (getCompanyByFio(emplMap, normalizeFio(name)) || '—') : '—');
  const withCompany = rows.map(r => ({ ...r, company: getCompany(r.name) }));
  const byCompany = new Map();
  for (const r of withCompany) {
    const c = r.company || '—';
    if (!byCompany.has(c)) byCompany.set(c, []);
    byCompany.get(c).push(r);
  }
  for (const arr of byCompany.values()) {
    arr.sort((a, b) => (b.total - a.total));
  }
  return { hours, byCompany: Object.fromEntries(byCompany) };
}

/** Стиль первой колонки (ФИО у левого края) — инлайн, чтобы html2canvas не терял при рендере */
const HE_NAME_COL_STYLE = 'width:200px;min-width:200px;max-width:200px;text-align:left;padding:6px 8px;border:1px solid #DDE2EA;background:#fff;font-weight:500;box-sizing:border-box;';

/**
 * HTML таблицы по часам для одной компании (для скриншота в Telegram).
 * ФИО — строго в первой колонке у левого края, часы — справа от неё.
 */
export function buildHourlyTableHtmlForCompany(companyName, rows, hours, dateStr, shiftLabel) {
  const hourLabel = (col) => {
    const start = (col + 23) % 24;
    return `${String(start).padStart(2, '0')}–${String(col).padStart(2, '0')}`;
  };
  const thHours = hours.map(col => `<th style="width:46px;padding:6px 8px;border:1px solid #DDE2EA;background:#f5f7fa;font-size:12px;text-align:center;" title="${hourLabel(col)}">${String(col).padStart(2, '0')}</th>`).join('');
  const thTotalStyle = 'width:56px;padding:6px 8px;border:1px solid #DDE2EA;background:#f5f7fa;font-size:12px;text-align:center;';
  const szCellClass = (v) => {
    if (v < 50) return 'he-sz-red';
    if (v <= 75) return 'he-sz-mid';
    return 'he-sz-white';
  };
  const trRows = rows.map(r => {
    const cells = hours.map(col => {
      const v = r.byHour[col] || 0;
      const cl = szCellClass(v);
      return `<td class="he-td-val ${cl}" style="width:46px;padding:6px 8px;border:1px solid #DDE2EA;text-align:center;">${v > 0 ? v : ''}</td>`;
    }).join('');
    const totalStyle = 'width:56px;padding:6px 8px;border:1px solid #DDE2EA;text-align:center;font-weight:600;';
    return `<tr><th scope="row" style="${HE_NAME_COL_STYLE}">${escHtml(r.name)}</th>${cells}<td style="${totalStyle}">${r.total}</td></tr>`;
  }).join('');
  return `
    <div class="he-telegram-wrap" style="padding:12px;background:#fff;font-family:Inter,sans-serif;">
      <div class="he-telegram-title" style="font-size:16px;font-weight:700;margin-bottom:4px;">${escHtml(companyName)}</div>
      <div class="he-telegram-meta" style="font-size:12px;color:#6b7280;margin-bottom:10px;">${escHtml(dateStr)} • ${escHtml(shiftLabel)}</div>
      <table style="border-collapse:collapse;table-layout:fixed;width:100%;font-size:13px;">
        <thead><tr><th style="${HE_NAME_COL_STYLE}background:#f5f7fa;font-size:12px;">Исполнитель</th>${thHours}<th style="${thTotalStyle}">Итого</th></tr></thead>
        <tbody>${trRows}</tbody>
      </table>
    </div>`;
}

/**
 * Рендерит таблицу «Сотрудник по часам». emplMap — для колонки «Компания».
 */
export function renderHourlyByEmployee(items, shiftFilter = 'day', emplMap = null) {
  const container = el('hourly-employee-table-wrap');
  if (!container) return;

  const { hours, rows } = calcHourlyByEmployee(items, shiftFilter);

  if (!rows.length) {
    container.innerHTML = '<div class="empty-row" style="padding:20px;text-align:center;color:var(--text-muted)">Нет данных</div>';
    return;
  }

  const getCompany = (name) => (emplMap && name ? (getCompanyByFio(emplMap, normalizeFio(name)) || '—') : '—');
  const withCompany = rows.map(r => ({ ...r, company: getCompany(r.name) }));
  withCompany.sort((a, b) => {
    if (b.total !== a.total) return b.total - a.total;
    return (b.company || '').localeCompare(a.company || '', 'ru');
  });

  const hourLabel = (col) => {
    const start = (col + 23) % 24;
    return `${String(start).padStart(2,'0')}–${String(col).padStart(2,'0')}`;
  };
  const thHours = hours.map(col => `<th class="he-th-hour" title="${hourLabel(col)}">${String(col).padStart(2,'0')}</th>`).join('');
  const szCellClass = (v) => {
    if (v < 50) return 'he-sz-red';
    if (v <= 75) return 'he-sz-mid';
    return 'he-sz-white';
  };

  const trRows = withCompany.map(r => {
    const cells = hours.map(col => {
      const v = r.byHour[col] || 0;
      const cl = szCellClass(v);
      return `<td class="he-td-val ${cl}" title="${hourLabel(col)} — ${v} оп.">${v > 0 ? v : ''}</td>`;
    }).join('');
    return `<tr>
      <td class="he-td-company">${escHtml(r.company)}</td>
      <td class="he-td-name">${escHtml(r.name)}</td>
      ${cells}
      <td class="he-td-total">${r.total}</td>
    </tr>`;
  }).join('');

  container.innerHTML = `
    <div class="he-scroll-wrap">
      <table class="he-table">
        <thead>
          <tr>
            <th class="he-th-company">Компания</th>
            <th class="he-th-name">Сотрудник</th>
            ${thHours}
            <th class="he-th-total">Итого</th>
          </tr>
        </thead>
        <tbody>${trRows}</tbody>
      </table>
    </div>
  `;
}

function escHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
