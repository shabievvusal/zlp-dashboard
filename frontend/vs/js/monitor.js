/**
 * monitor.js — мониторинг сотрудников по компаниям + перекличка
 *
 * Источник данных: GET /api/monitor/live
 *   value.pickByLineHandlingUnitsInProgress      — КДК в работе
 *   value.pieceSelectionHandlingUnitsInProgress  — ХР в работе
 *   value.palletSelectionHandlingUnitsInProgress — Паллеты в работе
 *
 * Каждая запись: { handlingUnitBarcode, startedAt, user: { id, firstName, lastName, middleName } }
 *
 * Логика:
 *  - "В работе"    = человек есть в live-ответе прямо сейчас
 *  - "Не работает" = человек есть в перекличке, но отсутствует в live-ответе
 *  - Таймер простоя: отсчитывается с момента исчезновения из live
 *  - При появлении: показываем сколько отсутствовал
 */

import { el, formatTime, normalizeFio } from './utils.js';
import * as api from './api.js';
import * as auth from './auth.js';

// ─── Состояние ───────────────────────────────────────────────────────────────

/** Set<normalizedFio> — кто отмечен на перекличке */
let rollcallPresent = new Set();
let rollcallShiftKey = null;

/**
 * Map<normalizedFio, { absentSince: number|null, wasAbsentMs: number|null }>
 * absentSince — timestamp когда зафиксировали исчезновение из live
 * wasAbsentMs — сколько мс отсутствовал до последнего возврата (для показа)
 */
const absentState = new Map();

/** Последний live-снапшот: Map<normalizedFio, { displayFio, taskType, startedAt }> */
let lastLiveSnapshot = new Map();

let monitorInterval = null;
let _emplMap = new Map();
let _emplCompanies = [];
/** Функция получения операций для расчёта времени от последней подтверждённой задачи */
let _getItemsFn = () => [];

/** Результаты точечных запросов за последние 30/60 мин (нормФИО → timestamp). Заполняется в refreshMonitor. */
let lastCompletedAtFromQueries = new Map();

/** Последний массив строк мониторинга (для клика по сотруднику). */
let _lastFlatRows = [];

/** Вызов после сохранения компании (обновить empl в приложении). */
let _onEmplSaved = () => {};
let _employeeModalSetup = false;

// ─── Инициализация ────────────────────────────────────────────────────────────

export function initMonitor(getItemsFn, emplMapRef, emplCompaniesRef, onEmplSaved) {
  _getItemsFn = getItemsFn || (() => []);
  _emplMap = emplMapRef;
  _emplCompanies = emplCompaniesRef;
  _onEmplSaved = onEmplSaved || (() => {});
  if (!_employeeModalSetup) {
    setupEmployeeModalListeners();
    _employeeModalSetup = true;
  }
}

export function updateMonitorEmpl(emplMap, emplCompanies) {
  _emplMap = emplMap;
  _emplCompanies = emplCompanies;
}

export async function loadRollcall() {
  try {
    const data = await api.getRollcall();
    rollcallPresent = new Set((data.present || []).map(f => normalizeFio(f)));
    rollcallShiftKey = data.shiftKey || null;
  } catch {
    rollcallPresent = new Set();
  }
}

export function getRollcallCount() {
  return rollcallPresent.size;
}

export function startMonitorRefresh() {
  if (monitorInterval) clearInterval(monitorInterval);
  monitorInterval = setInterval(() => refreshMonitor(), 10 * 60 * 1000);
}

export function stopMonitorRefresh() {
  if (monitorInterval) clearInterval(monitorInterval);
  monitorInterval = null;
}

// ─── Загрузка live-данных ─────────────────────────────────────────────────────

/**
 * Парсит ответ /api/monitor/live в Map<normalizedFio, { displayFio, taskType, startedAt }>
 */
function parseLiveData(data) {
  const result = new Map();
  const value = data?.value || data || {};

  const sections = [
    { key: 'pickByLineHandlingUnitsInProgress',     type: 'КДК' },
    { key: 'pieceSelectionHandlingUnitsInProgress',  type: 'ХР' },
    { key: 'palletSelectionHandlingUnitsInProgress', type: 'Паллет' },
  ];

  for (const { key, type } of sections) {
    for (const entry of (value[key] || [])) {
      const u = entry.user || {};
      const displayFio = [u.lastName, u.firstName, u.middleName].filter(Boolean).join(' ');
      if (!displayFio) continue;
      const normFio = normalizeFio(displayFio);
      // Один человек может вести только одну задачу — берём первую
      if (!result.has(normFio)) {
        result.set(normFio, {
          displayFio,
          userId: u.id || '',
          taskType: type,
          startedAt: entry.startedAt || null,
        });
      }
    }
  }
  return result;
}

/**
 * Обновляет absentState на основе нового снапшота.
 * Вызывается при каждом получении live-данных.
 */
function updateAbsentState(newSnapshot) {
  const now = Date.now();

  // Собираем personKey-множества для fuzzy-поиска по снапшотам
  const oldPKs = new Set([...lastLiveSnapshot.keys()].map(personKey));
  const newPKs = new Set([...newSnapshot.keys()].map(personKey));

  for (const normFio of rollcallPresent) {
    const pk = personKey(normFio);
    const wasInLive = lastLiveSnapshot.has(normFio) || oldPKs.has(pk);
    const isInLive  = newSnapshot.has(normFio) || newPKs.has(pk);
    let state = absentState.get(normFio);
    if (!state) { state = { absentSince: null, wasAbsentMs: null }; absentState.set(normFio, state); }

    if (isInLive) {
      // Сотрудник в работе — если до этого был absent, сохраняем длительность
      if (state.absentSince !== null) {
        state.wasAbsentMs = now - state.absentSince;
        state.absentSince = null;
      }
    } else {
      // Сотрудника нет в live
      if (wasInLive && state.absentSince === null) {
        // Только что пропал — фиксируем начало отсутствия
        state.absentSince = now;
        state.wasAbsentMs = null;
      } else if (!wasInLive && state.absentSince === null) {
        // Не было и раньше — фиксируем с момента первой проверки
        state.absentSince = now;
      }
    }
  }

  lastLiveSnapshot = newSnapshot;
}

// ─── Главная функция обновления ───────────────────────────────────────────────

export async function refreshMonitor() {
  const container = el('monitor-companies');
  const lastUpdEl = el('monitor-last-updated');
  const errorEl   = el('monitor-error');

  if (!container) return;

  // Индикатор загрузки
  if (errorEl) errorEl.style.display = 'none';

  let snapshot;
  try {
    const token = auth.getToken();
    const data = token
      ? await api.getLiveMonitorViaBrowser(token)
      : await api.getLiveMonitor();
    if (data && data.error) throw new Error(data.error);
    snapshot = parseLiveData(data || {});
  } catch (err) {
    if (errorEl) {
      errorEl.textContent = 'Ошибка загрузки live-данных: ' + err.message;
      errorEl.style.display = 'block';
    }
    // Рендерим с последним известным снапшотом
    snapshot = lastLiveSnapshot;
  }

  const token = auth.getToken();
  if (token && rollcallPresent.size > 0) {
    const executorIdByNorm = new Map();
    for (const [norm, entry] of snapshot) {
      if (entry.userId) executorIdByNorm.set(norm, entry.userId);
    }
    const items = _getItemsFn();
    for (const item of items) {
      if (!item.executorId) continue;
      const norm = normalizeFio(item.executor);
      if (!executorIdByNorm.has(norm)) executorIdByNorm.set(norm, item.executorId);
    }
    const now = new Date();
    const toIso = now.toISOString();
    const from30 = new Date(now.getTime() - 30 * 60 * 1000).toISOString();
    const from60 = new Date(now.getTime() - 60 * 60 * 1000).toISOString();
    lastCompletedAtFromQueries = new Map();
    for (const normFio of rollcallPresent) {
      const executorId = findByPersonKey(executorIdByNorm, normFio);
      if (!executorId) continue;
      try {
        let res = await api.fetchLastCompletedForExecutor(token, executorId, from30, toIso);
        if (res.maxCompletedAt == null) res = await api.fetchLastCompletedForExecutor(token, executorId, from60, toIso);
        if (res.maxCompletedAt != null) lastCompletedAtFromQueries.set(normFio, res.maxCompletedAt);
      } catch {
        // один сбой — не ломаем весь мониторинг
      }
    }
  } else {
    lastCompletedAtFromQueries = new Map();
  }

  updateAbsentState(snapshot);

  if (lastUpdEl) lastUpdEl.textContent = 'Обновлено: ' + formatTime(new Date().toISOString());

  renderMonitor();
}

// ─── Рендер ───────────────────────────────────────────────────────────────────

const expandedCompanies = new Set();

export function renderMonitor() {
  const container = el('monitor-companies');
  if (!container) return;

  if (rollcallPresent.size === 0) {
    const suggestionsHtml = renderSuggestionsSection();
    container.innerHTML = suggestionsHtml + `
      <div class="monitor-empty">
        Перекличка не проведена.<br>
        Нажмите <strong>«Перекличка»</strong> чтобы отметить сотрудников на смене.
      </div>`;
    attachSuggestionHandlers(container);
    return;
  }

  const now = Date.now();

  // Время от последней подтверждённой задачи: приоритет — точечные запросы (30/60 мин), иначе — из загруженных операций
  const items = _getItemsFn();
  const lastCompletedAtMap = new Map();
  for (const item of items) {
    const at = item.completedAt;
    if (!at) continue;
    const norm = normalizeFio(item.executor);
    const ts = new Date(at).getTime();
    if (!lastCompletedAtMap.has(norm) || lastCompletedAtMap.get(norm) < ts) {
      lastCompletedAtMap.set(norm, ts);
    }
  }
  for (const normFio of rollcallPresent) {
    const fromQuery = findByPersonKey(lastCompletedAtFromQueries, normFio);
    if (fromQuery != null) lastCompletedAtMap.set(normFio, fromQuery);
  }

  const IDLE_WORK_MIN = 10;

  // СЗ по сотрудникам (как в stats: ХР + КДК без двойного учёта)
  const szByNorm = new Map();
  for (const item of items) {
    const norm = normalizeFio(item.executor);
    if (!norm) continue;
    const type = (item.operationType || '').toUpperCase();
    const key = type === 'PICK_BY_LINE'
      ? `kdk|${item.executorId || ''}|${item.cell || ''}|${item.nomenclatureCode || item.productName || ''}`
      : (item.id ? `op|${item.id}` : `op|${item.completedAt || ''}|${item.executor || ''}|${item.cell || ''}`);
    if (!szByNorm.has(norm)) szByNorm.set(norm, new Set());
    szByNorm.get(norm).add(key);
  }

  // Один человек может быть в перекличке дважды (короткое и полное ФИО) — объединяем по «фамилия + имя»
  const byPersonKey = new Map();
  for (const normFio of rollcallPresent) {
    const key = personKey(normFio);
    if (!byPersonKey.has(key)) byPersonKey.set(key, []);
    byPersonKey.get(key).push(normFio);
  }
  const rollcallDeduped = [];
  for (const [, aliases] of byPersonKey) {
    const inLive = aliases.find(a => lastLiveSnapshot.has(a));
    const canonical = inLive || aliases.sort((a, b) => b.length - a.length)[0];
    rollcallDeduped.push({ canonical, aliases });
  }

  const rows = [];
  for (const { canonical, aliases } of rollcallDeduped) {
    const liveEntry = aliases.map(a => lastLiveSnapshot.get(a)).find(Boolean);
    const isActive = !!liveEntry;
    const company = findByPersonKey(_emplMap, canonical) || aliases.map(a => findByPersonKey(_emplMap, a)).find(Boolean) || '—';
    const state = absentState.get(canonical) || aliases.map(a => absentState.get(a)).find(Boolean) || { absentSince: null, wasAbsentMs: null };

    let lastTs = null;
    for (const a of aliases) {
      const t = lastCompletedAtMap.get(a);
      if (t != null && (lastTs == null || t > lastTs)) lastTs = t;
    }
    const minutesSinceLastTask = lastTs != null ? (now - lastTs) / 60000 : null;
    const lastTaskMs = lastTs != null ? now - lastTs : null;
    const inWorkByTask = minutesSinceLastTask != null && minutesSinceLastTask <= IDLE_WORK_MIN;

    const taskDurationMs = (isActive && liveEntry.startedAt)
      ? now - new Date(liveEntry.startedAt).getTime()
      : null;

    const idleMs = (!isActive && state.absentSince)
      ? now - state.absentSince
      : null;

    const wasAbsentMs = (isActive && state.wasAbsentMs && state.wasAbsentMs > 60000)
      ? state.wasAbsentMs
      : null;

    const displayFio = liveEntry?.displayFio || titleCase(canonical);
    const sz = aliases.reduce((s, a) => s + (szByNorm.get(a)?.size || 0), 0);

    rows.push({
      normFio: canonical,
      displayFio,
      company,
      isActive,
      aliases,
      sz,
      taskType: liveEntry?.taskType || null,
      startedAt: liveEntry?.startedAt || null,
      taskDurationMs,
      idleMs,
      wasAbsentMs,
      lastTaskMs,
      minutesSinceLastTask,
      inWorkByTask,
    });
  }

  // Группируем по компании
  const byCompany = new Map();
  for (const row of rows) {
    const c = row.company;
    if (!byCompany.has(c)) byCompany.set(c, { company: c, rows: [], active: 0, inactive: 0 });
    const g = byCompany.get(c);
    g.rows.push(row);
    if (row.isActive) g.active++; else g.inactive++;
  }

  // Сортировка: по времени от последней задачи (дольше простаивает — выше), потом по алфавиту
  const groups = [...byCompany.values()].sort((a, b) => {
    if (b.inactive !== a.inactive) return b.inactive - a.inactive;
    return a.company.localeCompare(b.company);
  });

  for (const g of groups) {
    g.rows.sort((a, b) => {
      const aMin = a.minutesSinceLastTask;
      const bMin = b.minutesSinceLastTask;
      if (aMin != null && bMin != null && aMin !== bMin) return bMin - aMin;
      if (aMin != null && bMin == null) return -1;
      if (aMin == null && bMin != null) return 1;
      return (b.idleMs || 0) - (a.idleMs || 0);
    });
  }

  const flatRows = groups.flatMap(g => g.rows);
  _lastFlatRows = flatRows;

  const suggestionsHtml = renderSuggestionsSection();
  const summaryHtml = renderOperationSummary(rows);
  container.innerHTML = suggestionsHtml + summaryHtml + groups.map(g => renderCompanyCard(g)).join('');
  attachSuggestionHandlers(container);

  container.onclick = (e) => {
    const tr = e.target.closest('.mon-row-clickable');
    if (!tr) return;
    e.preventDefault();
    e.stopPropagation();
    const normFio = tr.dataset.normFio;
    const row = _lastFlatRows.find(r => r.normFio === normFio);
    if (!row) return;
    openEmployeeModal({
      displayFio: row.displayFio,
      normFio: row.normFio,
      company: row.company,
      onShift: hasByPersonKey(rollcallPresent, row.normFio) || (row.aliases && row.aliases.some(a => hasByPersonKey(rollcallPresent, a))),
      aliases: row.aliases || [row.normFio],
    });
  };

  container.querySelectorAll('.mon-company-card').forEach(card => {
    card.querySelector('.mon-company-header').addEventListener('click', (e) => {
      e.stopPropagation();
      const c = card.dataset.company;
      if (expandedCompanies.has(c)) expandedCompanies.delete(c);
      else expandedCompanies.add(c);
      renderMonitor();
    });
  });
}

function renderOperationSummary(rows) {
  // Общие счётчики по типу операции
  const totalByType = { 'КДК': 0, 'ХР': 0, 'Паллет': 0 };
  // По компаниям: company → { КДК: N, ХР: N, Паллет: N }
  const byCompanyType = new Map();

  for (const r of rows) {
    if (!r.isActive || !r.taskType) continue;
    totalByType[r.taskType] = (totalByType[r.taskType] || 0) + 1;
    if (!byCompanyType.has(r.company)) byCompanyType.set(r.company, { 'КДК': 0, 'ХР': 0, 'Паллет': 0 });
    const ct = byCompanyType.get(r.company);
    ct[r.taskType] = (ct[r.taskType] || 0) + 1;
  }

  const totalActive = totalByType['КДК'] + totalByType['ХР'] + totalByType['Паллет'];
  if (totalActive === 0) return '';

  const typeBadge = (type, count) => {
    if (count === 0) return '';
    const cls = type === 'КДК' ? 'кдк' : type === 'ХР' ? 'хр' : 'паллет';
    return `<span class="mon-op-badge mon-task-type--${cls}">${escH(type)} <b>${count}</b></span>`;
  };

  // Строки по компаниям (только те, где есть активные)
  const companyRows = [...byCompanyType.entries()]
    .filter(([, ct]) => ct['КДК'] + ct['ХР'] + ct['Паллет'] > 0)
    .sort((a, b) => {
      const aSum = a[1]['КДК'] + a[1]['ХР'] + a[1]['Паллет'];
      const bSum = b[1]['КДК'] + b[1]['ХР'] + b[1]['Паллет'];
      return bSum - aSum;
    })
    .map(([company, ct]) => {
      const sum = ct['КДК'] + ct['ХР'] + ct['Паллет'];
      return `<div class="mon-op-company-row">
        <span class="mon-op-company-name">${escH(company)}</span>
        <span class="mon-op-company-total">${sum}</span>
        <span class="mon-op-badges">${typeBadge('КДК', ct['КДК'])}${typeBadge('ХР', ct['ХР'])}${typeBadge('Паллет', ct['Паллет'])}</span>
      </div>`;
    }).join('');

  return `
    <div class="mon-op-summary">
      <div class="mon-op-summary-header">
        <span class="mon-op-summary-title">В работе: <b>${totalActive}</b></span>
        <span class="mon-op-badges">
          ${typeBadge('КДК', totalByType['КДК'])}
          ${typeBadge('ХР', totalByType['ХР'])}
          ${typeBadge('Паллет', totalByType['Паллет'])}
        </span>
      </div>
      <div class="mon-op-company-list">${companyRows}</div>
    </div>`;
}

function renderCompanyCard(g) {
  const isExpanded = expandedCompanies.has(g.company);
  const total = g.active + g.inactive;
  const cardClass = g.inactive === 0
    ? 'mon-company-card mon-company-card--ok'
    : 'mon-company-card mon-company-card--warn';

  let expandedContent = '';
  if (isExpanded) {
    const activeRows = g.rows.filter(r => r.isActive || r.inWorkByTask);
    const inactiveRows = g.rows.filter(r => !r.isActive && !r.inWorkByTask);

    const tableHead = `<thead><tr>
      <th>ФИО</th><th>Тип задачи</th><th>Статус</th><th>В задаче / Простой</th><th class="mon-th-sz">СЗ</th>
    </tr></thead>`;

    const activeHtml = activeRows.length
      ? `<table class="mon-emp-table">${tableHead}<tbody>${activeRows.map(r => renderEmployeeRow(r)).join('')}</tbody></table>`
      : '<div class="mon-col-empty">Нет активных</div>';

    const inactiveHtml = inactiveRows.length
      ? `<table class="mon-emp-table">${tableHead}<tbody>${inactiveRows.map(r => renderEmployeeRow(r)).join('')}</tbody></table>`
      : '<div class="mon-col-empty">Все в работе</div>';

    expandedContent = `
      <div class="mon-employee-list">
        <div class="mon-two-col">
          <div class="mon-col mon-col--active">
            <div class="mon-col-header mon-col-header--active">В работе (${activeRows.length})</div>
            ${activeHtml}
          </div>
          <div class="mon-col mon-col--inactive">
            <div class="mon-col-header mon-col-header--inactive">Не работают (${inactiveRows.length})</div>
            ${inactiveHtml}
          </div>
        </div>
      </div>`;
  }

  return `
    <div class="${cardClass}" data-company="${escH(g.company)}">
      <div class="mon-company-header">
        <div class="mon-company-title">
          <span class="mon-company-name">${escH(g.company)}</span>
          <span class="mon-company-count">${total} чел.</span>
        </div>
        <div class="mon-company-badges">
          <span class="mon-badge mon-badge--active">${g.active} в работе</span>
          ${g.inactive > 0 ? `<span class="mon-badge mon-badge--inactive">${g.inactive} не работают</span>` : ''}
        </div>
        <span class="mon-expand-icon">${isExpanded ? '▲' : '▼'}</span>
      </div>
      ${expandedContent}
    </div>`;
}

function renderEmployeeRow(r) {
  const wasAbsent = r.wasAbsentMs
    ? `<span class="mon-was-absent">вернулся (был ${formatDuration(r.wasAbsentMs)})</span>`
    : '';
  const szVal = r.sz != null && r.sz > 0 ? String(r.sz) : '—';

  if (r.isActive) {
    // В работе: показываем длительность текущей задачи из live startedAt
    const taskStr = r.taskDurationMs ? formatDuration(r.taskDurationMs) : '—';
    // В скобках — время от последней завершённой задачи (простой между задачами)
    const lastTaskHint = r.lastTaskMs != null ? ` <span class="mon-idle-hint">(посл. ${formatDuration(r.lastTaskMs)} назад)</span>` : '';
    return `
      <tr class="mon-row--active mon-row-clickable" data-norm-fio="${escH(r.normFio)}" title="Нажмите, чтобы изменить компанию и статус на смене">
        <td class="mon-td-name">${escH(r.displayFio)}</td>
        <td class="mon-td-type"><span class="mon-task-type mon-task-type--${(r.taskType||'').toLowerCase()}">${escH(r.taskType || '—')}</span></td>
        <td class="mon-td-status">🟢 в работе ${wasAbsent}</td>
        <td class="mon-td-idle"><span class="mon-idle-time mon-idle-time--work">${taskStr}</span>${lastTaskHint}</td>
        <td class="mon-td-sz">${szVal}</td>
      </tr>`;
  } else {
    // Не в работе: показываем время от последней завершённой задачи, в скобках — простой из absentState
    const lastTaskStr = r.lastTaskMs != null ? formatDuration(r.lastTaskMs) : '—';
    const idleClass = r.inWorkByTask ? 'mon-idle-time--work' : 'mon-idle-time--idle';
    const idleHint = r.idleMs != null ? ` <span class="mon-idle-hint">(нет в live ${formatDuration(r.idleMs)})</span>` : '';
    const statusByTask = r.inWorkByTask ? '🟢 в работе' : '🔴 в простое';
    return `
      <tr class="mon-row--inactive mon-row-clickable" data-norm-fio="${escH(r.normFio)}" title="Нажмите, чтобы изменить компанию и статус на смене">
        <td class="mon-td-name">${escH(r.displayFio)}</td>
        <td class="mon-td-type">—</td>
        <td class="mon-td-status">${statusByTask}</td>
        <td class="mon-td-idle"><span class="mon-idle-time ${idleClass}">${lastTaskStr}</span>${idleHint}</td>
        <td class="mon-td-sz">${szVal}</td>
      </tr>`;
  }
}

function formatDuration(ms) {
  if (!ms || ms < 0) return '—';
  const totalMin = Math.floor(ms / 60000);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  if (h > 0) return `${h} ч ${m} мин`;
  return `${m} мин`;
}

// ─── Модальное окно редактирования сотрудника ───────────────────────────────────

let _employeeModalRow = null;

export function openEmployeeModal(row) {
  _employeeModalRow = row;
  const modal = el('employee-edit-modal');
  const fioEl = el('employee-edit-fio');
  const companySel = el('employee-edit-company');
  const onshiftCb = el('employee-edit-onshift');
  if (!modal || !fioEl || !companySel || !onshiftCb) return;

  fioEl.textContent = row.displayFio || row.normFio || '—';
  const companies = ['', ..._emplCompanies];
  const currentCompany = row.company === '—' ? '' : (row.company || '');
  companySel.innerHTML = companies.map(c => `<option value="${escH(c)}"${c === currentCompany ? ' selected' : ''}>${escH(c || '— не указана —')}</option>`).join('');
  onshiftCb.checked = !!row.onShift;

  modal.classList.add('modal--open');
  modal.addEventListener('click', e => { if (e.target === modal) closeEmployeeModal(); }, { once: true });
}

export function closeEmployeeModal() {
  _employeeModalRow = null;
  const modal = el('employee-edit-modal');
  if (modal) modal.classList.remove('modal--open');
}

function setupEmployeeModalListeners() {
  el('btn-employee-modal-close')?.addEventListener('click', closeEmployeeModal);
  el('btn-employee-modal-cancel')?.addEventListener('click', closeEmployeeModal);
  el('btn-employee-modal-save')?.addEventListener('click', saveEmployeeModal);
}

async function saveEmployeeModal() {
  const row = _employeeModalRow;
  if (!row) { closeEmployeeModal(); return; }

  const companySel = el('employee-edit-company');
  const onshiftCb = el('employee-edit-onshift');
  const newCompany = (companySel?.value || '').trim() || '—';
  const newOnShift = onshiftCb?.checked ?? false;

  let companyChanged = (row.company === '—' ? '' : row.company) !== (newCompany === '—' ? '' : newCompany);
  if (companyChanged) {
    try {
      await api.saveEmplOne(row.displayFio || row.normFio, newCompany === '—' ? '' : newCompany);
      _onEmplSaved();
    } catch (e) {
      console.error('Сохранение компании', e);
    }
  }

  const aliases = row.aliases || [row.normFio];
  const wasOnShift = hasByPersonKey(rollcallPresent, row.normFio) || aliases.some(a => hasByPersonKey(rollcallPresent, a));
  if (newOnShift !== wasOnShift) {
    let present = [...rollcallPresent];
    if (newOnShift) {
      for (const a of aliases) { if (!present.includes(a)) present.push(a); }
    } else {
      present = present.filter(p => !aliases.includes(p));
    }
    rollcallPresent = new Set(present);
    try {
      await api.putRollcall(rollcallShiftKey, present);
    } catch { /* локально уже обновили */ }
    await loadRollcall();
  }

  closeEmployeeModal();
  renderMonitor();
}

// ─── Модальное окно переклички ────────────────────────────────────────────────

export function openRollcallModal(currentShiftKey) {
  const modal = el('rollcall-modal');
  if (!modal) return;
  renderRollcallModal(currentShiftKey);
  modal.classList.add('modal--open');
  modal.addEventListener('click', e => {
    if (e.target === modal) closeRollcallModal();
  }, { once: true });
}

export function closeRollcallModal() {
  const modal = el('rollcall-modal');
  if (modal) modal.classList.remove('modal--open');
}

function renderRollcallModal() {
  const body = el('rollcall-modal-body');
  if (!body) return;

  if (!_emplMap.size && !lastLiveSnapshot.size) {
    body.innerHTML = '<p style="color:var(--text-muted)">Список сотрудников пуст. Добавьте сотрудников в настройках.</p>';
    return;
  }

  // Собираем все ФИО из empl + live
  const allFios = new Map(); // normFio -> { displayFio, company }
  for (const [normFio, company] of _emplMap) {
    allFios.set(normFio, { displayFio: titleCase(normFio), company: company || '—' });
  }
  for (const [normFio, entry] of lastLiveSnapshot) {
    if (!allFios.has(normFio)) {
      allFios.set(normFio, { displayFio: entry.displayFio, company: '—' });
    }
  }

  // Дедуплицируем по personKey: один чекбокс на группу алиасов
  const byPK = new Map(); // personKey -> [{ normFio, displayFio, company }]
  for (const [normFio, info] of allFios) {
    const pk = personKey(normFio);
    if (!byPK.has(pk)) byPK.set(pk, []);
    byPK.get(pk).push({ normFio, ...info });
  }

  const deduped = []; // { canonical, displayFio, company, aliases }
  for (const [, group] of byPK) {
    // Самое длинное ФИО = каноническое
    group.sort((a, b) => b.normFio.length - a.normFio.length);
    const canonical = group[0].normFio;
    const displayFio = group[0].displayFio;
    // Предпочесть реальную компанию (не «—»)
    const company = group.find(g => g.company !== '—')?.company || group[0].company;
    const aliases = group.map(g => g.normFio);
    const isChecked = aliases.some(a => rollcallPresent.has(a));
    deduped.push({ canonical, displayFio, company, aliases, isChecked });
  }

  // Группируем по компании
  const groups = new Map();
  for (const item of deduped) {
    if (!groups.has(item.company)) groups.set(item.company, []);
    groups.get(item.company).push(item);
  }

  const sorted = [...groups.entries()].sort((a, b) => a[0].localeCompare(b[0]));

  body.innerHTML = sorted.map(([company, items]) => `
    <div class="rc-group">
      <div class="rc-group-header">
        <span class="rc-group-name">${escH(company)}</span>
        <button class="rc-btn-all" data-company="${escH(company)}">Все</button>
        <button class="rc-btn-none" data-company="${escH(company)}">Никого</button>
      </div>
      <div class="rc-group-rows">
        ${items.sort((a,b) => a.displayFio.localeCompare(b.displayFio)).map(item => {
          const checked = item.isChecked ? 'checked' : '';
          return `<label class="rc-row">
            <input type="checkbox" class="rc-check" data-fio="${escH(item.canonical)}" ${checked}>
            <span>${escH(item.displayFio)}</span>
          </label>`;
        }).join('')}
      </div>
    </div>
  `).join('');

  body.querySelectorAll('.rc-btn-all').forEach(btn => {
    btn.addEventListener('click', () => {
      const c = btn.dataset.company;
      body.querySelectorAll('.rc-group').forEach(g => {
        if (g.querySelector('.rc-group-name')?.textContent === c) {
          g.querySelectorAll('.rc-check').forEach(cb => cb.checked = true);
        }
      });
    });
  });

  body.querySelectorAll('.rc-btn-none').forEach(btn => {
    btn.addEventListener('click', () => {
      const c = btn.dataset.company;
      body.querySelectorAll('.rc-group').forEach(g => {
        if (g.querySelector('.rc-group-name')?.textContent === c) {
          g.querySelectorAll('.rc-check').forEach(cb => cb.checked = false);
        }
      });
    });
  });
}

export async function saveRollcall(shiftKey) {
  const body = el('rollcall-modal-body');
  if (!body) return;

  const present = [];
  body.querySelectorAll('.rc-check:checked').forEach(cb => {
    present.push(cb.dataset.fio);
  });

  rollcallPresent = new Set(present);
  rollcallShiftKey = shiftKey;

  // Сброс таймеров для тех кого убрали из переклички
  for (const [normFio] of absentState) {
    if (!hasByPersonKey(rollcallPresent, normFio)) absentState.delete(normFio);
  }

  try {
    await api.putRollcall(shiftKey, present);
  } catch { /* работаем локально */ }

  closeRollcallModal();
  renderMonitor();
}

// ─── Подсказки: добавить в перекличку ────────────────────────────────────────

/** Возвращает HTML-блок с людьми из live, отсутствующими в перекличке */
function renderSuggestionsSection() {
  if (!lastLiveSnapshot.size) return '';

  const suggestions = [];
  const seenPKs = new Set();

  // Собираем personKey всех кто уже в перекличке
  const rollcallPKs = new Set([...rollcallPresent].map(personKey));

  for (const [normFio, entry] of lastLiveSnapshot) {
    const pk = personKey(normFio);
    if (rollcallPKs.has(pk) || seenPKs.has(pk)) continue;
    seenPKs.add(pk);
    const company = findByPersonKey(_emplMap, normFio) || '—';
    suggestions.push({ normFio, displayFio: entry.displayFio, company, taskType: entry.taskType });
  }

  if (!suggestions.length) return '';

  const items = suggestions.map(s => `
    <div class="mon-suggestion-item">
      <span class="mon-suggestion-name">${escH(s.displayFio)}</span>
      <span class="mon-suggestion-company">${escH(s.company)}</span>
      <span class="mon-suggestion-task">${escH(s.taskType || '')}</span>
      <button class="mon-suggestion-add" data-fio="${escH(s.normFio)}">+ Добавить</button>
    </div>
  `).join('');

  return `
    <div class="mon-suggestions">
      <div class="mon-suggestions-header">
        Обнаружены в системе, но не в перекличке (${suggestions.length})
      </div>
      <div class="mon-suggestions-list">${items}</div>
      <div class="mon-suggestions-footer">
        <button class="mon-suggestion-add-all">Добавить всех (${suggestions.length})</button>
      </div>
    </div>`;
}

/** Добавляет человека в перекличку и перерендеривает */
async function addToRollcall(normFio) {
  const present = [...rollcallPresent];
  if (!present.includes(normFio)) present.push(normFio);
  rollcallPresent = new Set(present);
  try {
    await api.putRollcall(rollcallShiftKey, present);
  } catch { /* локально уже обновили */ }
  renderMonitor();
}

/** Вешает обработчики на кнопки подсказок */
function attachSuggestionHandlers(container) {
  container.querySelectorAll('.mon-suggestion-add').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      addToRollcall(btn.dataset.fio);
    });
  });

  const addAllBtn = container.querySelector('.mon-suggestion-add-all');
  if (addAllBtn) {
    addAllBtn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const btns = container.querySelectorAll('.mon-suggestion-add');
      const present = [...rollcallPresent];
      btns.forEach(b => {
        if (!present.includes(b.dataset.fio)) present.push(b.dataset.fio);
      });
      rollcallPresent = new Set(present);
      try {
        await api.putRollcall(rollcallShiftKey, present);
      } catch { /* локально уже обновили */ }
      renderMonitor();
    });
  }
}

// ─── Утилиты ─────────────────────────────────────────────────────────────────

/** Извлекает первые 2 слова (фамилия + имя) для fuzzy-матчинга ФИО */
function personKey(norm) {
  return norm.split(/\s+/).slice(0, 2).join(' ') || norm;
}

/** Ищет в Map: сначала точный ключ, затем fallback по personKey */
function findByPersonKey(map, normFio) {
  if (map.has(normFio)) return map.get(normFio);
  const pk = personKey(normFio);
  for (const [k, v] of map) {
    if (personKey(k) === pk) return v;
  }
  return undefined;
}

/** Ищет в Set: сначала точный, затем по personKey */
function hasByPersonKey(set, normFio) {
  if (set.has(normFio)) return true;
  const pk = personKey(normFio);
  for (const v of set) {
    if (personKey(v) === pk) return true;
  }
  return false;
}

function titleCase(s) {
  return (s || '').split(' ').map(w => w ? w[0].toUpperCase() + w.slice(1) : '').join(' ');
}

function escH(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
