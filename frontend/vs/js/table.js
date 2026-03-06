/**
 * table.js — таблица операций с пагинацией, поиском и сортировкой
 */

import { el, formatDateTime, exportToCsv, normalizeFio, getCompanyByFio } from './utils.js';

const PAGE_SIZE = 100;
let currentPage = 1;
let allRows = [];
let filteredRows = [];
let searchQuery = '';
let sortCol = 'completedAt';
let sortDir = -1; // -1 desc, 1 asc
let emplMap = null;

const COL_LABELS = {
  id: 'ID',
  type: 'Тип',
  operationType: 'Тип операции',
  productName: 'Товар',
  nomenclatureCode: 'Код номенкл.',
  barcodes: 'Штрихкоды',
  productionDate: 'Дата пр-ва',
  bestBeforeDate: 'Годен до',
  sourceBarcode: 'ШК источника',
  cell: 'Ячейка',
  targetBarcode: 'ШК приёмника',
  startedAt: 'Начало',
  completedAt: 'Окончание',
  company: 'Компания',
  executor: 'Ответственный',
  srcOld: 'Ист. было',
  srcNew: 'Ист. стало',
  tgtOld: 'Приём было',
  tgtNew: 'Приём стало',
};

export function setTableData(rows, emplMapArg = null) {
  allRows = rows;
  emplMap = emplMapArg;
  currentPage = 1;
  applyFilters();
}

export function setSearch(query) {
  searchQuery = query.toLowerCase();
  currentPage = 1;
  applyFilters();
}

function applyFilters() {
  if (!searchQuery) {
    filteredRows = [...allRows];
  } else {
    filteredRows = allRows.filter(r =>
      (r.executor || '').toLowerCase().includes(searchQuery) ||
      (r.productName || '').toLowerCase().includes(searchQuery) ||
      (r.nomenclatureCode || '').toLowerCase().includes(searchQuery) ||
      (r.cell || '').toLowerCase().includes(searchQuery) ||
      (r.barcodes || '').toLowerCase().includes(searchQuery) ||
      (r.id || '').toLowerCase().includes(searchQuery) ||
      (r.sourceBarcode || '').toLowerCase().includes(searchQuery) ||
      (r.targetBarcode || '').toLowerCase().includes(searchQuery)
    );
  }
  sortRows();
  renderTable();
  renderPagination();
}

const DATE_COLS = new Set(['completedAt', 'startedAt', 'productionDate', 'bestBeforeDate']);
const NUM_COLS = new Set(['srcOld', 'srcNew', 'tgtOld', 'tgtNew']);

function sortRows() {
  const getCompany = (r) => (emplMap && r.executor ? (getCompanyByFio(emplMap, normalizeFio(r.executor)) || '') : '');
  filteredRows.sort((a, b) => {
    const va = sortCol === 'company' ? getCompany(a) : (a[sortCol] ?? '');
    const vb = sortCol === 'company' ? getCompany(b) : (b[sortCol] ?? '');
    if (DATE_COLS.has(sortCol)) {
      return sortDir * (new Date(va || 0) - new Date(vb || 0));
    }
    if (NUM_COLS.has(sortCol)) {
      return sortDir * ((Number(va) || 0) - (Number(vb) || 0));
    }
    return sortDir * String(va).localeCompare(String(vb), 'ru');
  });
}

export function renderTable() {
  const tbody = el('ops-tbody');
  if (!tbody) return;

  const start = (currentPage - 1) * PAGE_SIZE;
  const page = filteredRows.slice(start, start + PAGE_SIZE);

  const colCount = Object.keys(COL_LABELS).length;
  if (!page.length) {
    tbody.innerHTML = `<tr><td colspan="${colCount}" class="empty-row">Нет данных для отображения</td></tr>`;
    return;
  }

  const getCompany = (r) => (emplMap && r.executor ? (getCompanyByFio(emplMap, normalizeFio(r.executor)) || '—') : '—');

  const v = (val) => (val != null && val !== '' && val !== undefined) ? escHtml(String(val)) : '—';

  tbody.innerHTML = page.map(r => `
    <tr>
      <td class="td-code" title="${escHtml(r.id || '')}">${escHtml(truncate(r.id || '', 8))}</td>
      <td class="td-type">${v(r.type)}</td>
      <td class="td-type">${formatOpType(r.operationType)}</td>
      <td class="td-product" title="${escHtml(r.productName || '')}">${escHtml(truncate(r.productName || '', 40))}</td>
      <td class="td-code">${v(r.nomenclatureCode)}</td>
      <td class="td-code">${v(r.barcodes)}</td>
      <td class="td-time">${v(r.productionDate)}</td>
      <td class="td-time">${v(r.bestBeforeDate)}</td>
      <td class="td-code">${v(r.sourceBarcode)}</td>
      <td class="td-cell">${v(r.cell)}</td>
      <td class="td-code">${v(r.targetBarcode)}</td>
      <td class="td-time">${formatDateTime(r.startedAt)}</td>
      <td class="td-time">${formatDateTime(r.completedAt)}</td>
      <td class="td-company">${escHtml(getCompany(r))}</td>
      <td class="td-executor">${v(r.executor)}</td>
      <td class="td-qty text-right">${v(r.srcOld)}</td>
      <td class="td-qty text-right">${v(r.srcNew)}</td>
      <td class="td-qty text-right">${v(r.tgtOld)}</td>
      <td class="td-qty text-right">${v(r.tgtNew)}</td>
    </tr>
  `).join('');

  const counter = el('table-counter');
  if (counter) counter.textContent = `Показано ${start + 1}–${Math.min(start + PAGE_SIZE, filteredRows.length)} из ${filteredRows.length}`;
}

function renderPagination() {
  const wrap = el('pagination');
  if (!wrap) return;

  const totalPages = Math.ceil(filteredRows.length / PAGE_SIZE);
  if (totalPages <= 1) { wrap.innerHTML = ''; return; }

  let html = '';
  if (currentPage > 1) html += `<button class="page-btn" data-page="${currentPage - 1}">‹</button>`;

  const start = Math.max(1, currentPage - 2);
  const end = Math.min(totalPages, currentPage + 2);
  for (let p = start; p <= end; p++) {
    html += `<button class="page-btn${p === currentPage ? ' active' : ''}" data-page="${p}">${p}</button>`;
  }

  if (currentPage < totalPages) html += `<button class="page-btn" data-page="${currentPage + 1}">›</button>`;
  wrap.innerHTML = html;

  wrap.querySelectorAll('.page-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      currentPage = parseInt(btn.dataset.page, 10);
      renderTable();
      renderPagination();
    });
  });
}

export function initTableHeaders() {
  const thead = el('ops-thead');
  if (!thead) return;

  thead.innerHTML = Object.entries(COL_LABELS).map(([col, label]) => `
    <th class="th-sortable${sortCol === col ? ' th-sorted' : ''}" data-col="${col}">
      ${label}${sortCol === col ? (sortDir === -1 ? ' ↓' : ' ↑') : ''}
    </th>
  `).join('');

  thead.querySelectorAll('.th-sortable').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (sortCol === col) sortDir *= -1;
      else { sortCol = col; sortDir = -1; }
      initTableHeaders();
      sortRows();
      currentPage = 1;
      renderTable();
      renderPagination();
    });
  });
}

export function exportTable(filename) {
  exportToCsv(filteredRows, filename || 'operations.csv');
}

function truncate(str, n) {
  if (!str) return '';
  return str.length > n ? str.slice(0, n) + '…' : str;
}

function escHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function formatOpType(t) {
  if (!t) return '—';
  if (t.includes('PIECE_SELECTION')) return 'Штучный';
  if (t.includes('PICK_BY_LINE')) return 'По линии';
  return t;
}
