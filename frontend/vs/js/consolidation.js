/**
 * consolidation.js — модуль вкладки «Консолидация»
 * WMS-поиск выполняется из браузера с токеном пользователя
 */

import {
  getConsolidationComplaints,
  updateComplaintStatus,
  deleteComplaint,
  saveComplaintLookup,
  sendComplaintsToTelegram,
} from './api.js';
import { getToken } from './auth.js';

const SAMOKAT_STOCKS_URL = 'https://api.samokat.ru/wmsops-wwh/stocks/changes/search';
const SAMOKAT_CELLS_BY_ADDRESS_URL = 'https://api.samokat.ru/wmsops-wwh/topology/cells/filters/by-address-search';
const LOOKUP_OPERATION_TYPES = [
  'PIECE_SELECTION_PICKING',
  'PIECE_SELECTION_PICKING_COMPLETE',
  'PICK_BY_LINE',
  'PALLET_SELECTION_MOVE_TO_PICK_BY_LINE',
];

let allComplaints = [];
let statusFilter = 'all';
let selectedComplaintIds = new Set();
let modalPhotoUrls = [];
let modalPhotoIndex = 0;
let modalControlsBound = false;

function esc(s) {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function formatDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  const pad = n => String(n).padStart(2, '0');
  return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function statusLabel(status) {
  const map = { new: 'Новая', in_progress: 'В работе', resolved: 'Решена' };
  return map[status] || status;
}

function statusClass(status) {
  const map = { new: 'cons-status--new', in_progress: 'cons-status--progress', resolved: 'cons-status--resolved' };
  return map[status] || '';
}

// ─── WMS-поиск из браузера ──────────────────────────────────────────────────

async function wmsPost(token, body) {
  const r = await fetch(SAMOKAT_STOCKS_URL, {
    method: 'POST',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
      'Origin': 'https://wwh.samokat.ru',
      'Referer': 'https://wwh.samokat.ru/',
    },
    body: JSON.stringify(body),
  });
  const text = await r.text();
  if (!r.ok) throw new Error(`API ${r.status}`);
  return JSON.parse(text);
}

async function wmsGet(token, url) {
  const r = await fetch(url, {
    method: 'GET',
    headers: {
      'Accept': 'application/json',
      'Authorization': `Bearer ${token}`,
      'Origin': 'https://wwh.samokat.ru',
      'Referer': 'https://wwh.samokat.ru/',
    },
  });
  const text = await r.text();
  if (!r.ok) throw new Error(`API ${r.status}`);
  return JSON.parse(text);
}

function normalizeCellAddress(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[—–−]/g, '-');
}

function extractCellIdList(data, wantedAddressNorm) {
  const value = data?.value || data || {};
  const candidateLists = [
    value?.items,
    data?.items,
    value?.content,
    data?.content,
    value?.cells,
    data?.cells,
    Array.isArray(value) ? value : null,
    Array.isArray(data) ? data : null,
  ].filter(Array.isArray);

  const rawItems = candidateLists.flat();
  const outExact = [];
  const outLoose = [];
  const wantedNorm = normalizeCellAddress(wantedAddressNorm);

  for (const it of rawItems) {
    const id = it?.cellId ?? it?.id ?? null;
    const addr = normalizeCellAddress(
      it?.cellAddress || it?.fullAddress || it?.address || it?.name || ''
    );
    if (!id) continue;

    if (wantedNorm) {
      if (addr === wantedNorm) {
        if (!outExact.includes(id)) outExact.push(id);
      } else if (addr.includes(wantedNorm) || wantedNorm.includes(addr)) {
        if (!outLoose.includes(id)) outLoose.push(id);
      }
    } else if (!outLoose.includes(id)) {
      outLoose.push(id);
    }
  }

  if (outExact.length > 0) return outExact;
  return outLoose;
}

async function findCellIdsByAddress(token, cellAddress) {
  const query = String(cellAddress || '').trim();
  if (!query) return [];
  const urls = [
    `${SAMOKAT_CELLS_BY_ADDRESS_URL}?cellAddressSearch=${encodeURIComponent(query)}`,
    `${SAMOKAT_CELLS_BY_ADDRESS_URL}?cellAddressSearch=${encodeURIComponent(query)}&pageNumber=1&pageSize=50`,
  ];

  for (const url of urls) {
    try {
      const data = await wmsGet(token, url);
      const exactOrLoose = extractCellIdList(data, query);
      if (exactOrLoose.length > 0) return exactOrLoose;

      const any = extractCellIdList(data, '');
      if (any.length > 0) return any;
    } catch {
      // пробуем следующий вариант запроса
    }
  }

  return [];
}

function fioFromUser(user) {
  if (!user) return null;
  if (typeof user === 'string') return user;
  return [user.lastName, user.firstName, user.middleName].filter(Boolean).join(' ').trim() || null;
}

function matchesBarcode(item, barcodeNorm) {
  const barcodes = item?.product?.barcodes || [];
  const barcodeMatch = barcodes.some(b => String(b).trim() === barcodeNorm);
  const nomenclatureMatch = String(item?.product?.nomenclatureCode || '').trim() === barcodeNorm;
  return barcodeMatch || nomenclatureMatch;
}

function matchesHandlingUnitBarcode(item, barcodeNorm) {
  const srcHu = String(item?.sourceAddress?.handlingUnitBarcode || '').trim();
  const tgtHu = String(item?.targetAddress?.handlingUnitBarcode || '').trim();
  return srcHu === barcodeNorm || tgtHu === barcodeNorm;
}

function pickProductBarcode(item, requestedBarcode) {
  const list = Array.isArray(item?.product?.barcodes) ? item.product.barcodes.map(x => String(x).trim()).filter(Boolean) : [];
  const req = String(requestedBarcode || '').trim();
  if (req && list.includes(req)) return req;
  return list[0] || null;
}

function matchesCell(item, cellNorm) {
  const target = String(item?.targetAddress?.cellAddress || '').trim().toLowerCase();
  const source = String(item?.sourceAddress?.cellAddress || '').trim().toLowerCase();
  return target === cellNorm || source === cellNorm;
}

function matchesTargetCell(item, cellIds, cellNorm) {
  if (!Array.isArray(cellIds) || cellIds.length === 0) {
    return matchesCell(item, cellNorm);
  }
  const targetCellId = item?.targetAddress?.cellId;
  if (targetCellId) return cellIds.includes(targetCellId);
  // В некоторых ответах WMS targetAddress.cellId отсутствует, остаётся только cellAddress.
  return matchesCell(item, cellNorm);
}

/**
 * Поиск продукта + нарушителя из браузера.
 *
 * Шаг A: ищем операции по месту (cellId), затем на клиенте фильтруем по barcode товара.
 *   Явный фильтр по barcode в payload не используем.
 *
 * Шаг B: по productId + ячейке ищем кто последний работал с этим товаром в этой ячейке.
 */
async function lookupViaBrowser(token, barcode, cell) {
  // ─── Вспомогательная функция ISO МСК ─────────────
  function isoMsk(date) {
    const tzOffset = -3 * 60; // Москва UTC+3, минус для смещения
    const local = new Date(date.getTime() - tzOffset * 60000);
    return local.toISOString().replace('Z', '+03:00');
  }

  // ─── Сегодняшнее начало дня и текущее время МСК ───
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayISO = isoMsk(today);
  const nowISO = isoMsk(new Date());

  const result = {
    productName: null,
    nomenclatureCode: null,
    productBarcode: null,
    violator: null,
    violatorId: null,
    handlingUnitBarcode: null,
    operationType: null,
    operationCompletedAt: null,
    lookupDone: true,
    lookupError: null,
    strategy: null,
  };

  const barcodeNorm = String(barcode).trim();
  const cellNorm = String(cell || '').trim().toLowerCase();
  let cellIds = [];
  let strategy = null;
  let itemsA = [];

  try {
    cellIds = await findCellIdsByAddress(token, cell);
    if (cellIds.length > 0) {
      console.log(`[Консолидация] Ячейка "${cell}" -> cellId: ${cellIds.join(', ')}`);
    } else {
      console.log(`[Консолидация] Ячейка "${cell}" не резолвится в cellId, fallback по адресу`);
    }
  } catch (e) {
    console.log(`[Консолидация] Ошибка резолва ячейки "${cell}": ${e.message}`);
  }

  const baseBody = {
    productId: null,
    parts: [],
    operationTypes: null,
    sourceCellId: null,
    targetCellId: null,
    operationStartedAtFrom: todayISO,
    operationStartedAtTo: nowISO,
    operationCompletedAtFrom: todayISO,
    operationCompletedAtTo: nowISO,
    executorId: null,
  };

  // ─── Приоритетный путь: точный матч (ячейка + штрихкод) ─────────────
  // API не умеет фильтровать по EAN и адресу напрямую, поэтому идем по страницам
  // и фильтруем строго по двум условиям на клиенте.
  try {
    const exactBodies = [];
    if (cellIds.length > 0) {
      for (const id of cellIds) {
        exactBodies.push({
          ...baseBody,
          targetCellId: id,
          operationTypes: LOOKUP_OPERATION_TYPES,
        });
      }
    } else {
      exactBodies.push({
        ...baseBody,
        operationTypes: LOOKUP_OPERATION_TYPES,
      });
    }

    const pageSize = 500;
    let pageNumber = 1;
    let exactFound = [];
    let exactMatchMode = null;
    while (true) {
      const batches = await Promise.all(
        exactBodies.map(body => wmsPost(token, { ...body, pageNumber, pageSize }))
      );
      const allItems = batches.flatMap(b => b?.value?.items || []);
      if (allItems.length === 0) break;

      const byProductBarcode = allItems.filter(it => {
        return matchesTargetCell(it, cellIds, cellNorm) && matchesBarcode(it, barcodeNorm);
      });
      if (byProductBarcode.length > 0) {
        exactFound = byProductBarcode;
        exactMatchMode = 'product_barcode';
        break;
      }

      const byHandlingUnit = allItems.filter(it => {
        return matchesTargetCell(it, cellIds, cellNorm) && matchesHandlingUnitBarcode(it, barcodeNorm);
      });
      if (byHandlingUnit.length > 0) {
        exactFound = byHandlingUnit;
        exactMatchMode = 'handling_unit_barcode';
        break;
      }

      if (exactFound.length > 0) break;
      pageNumber++;
    }

    if (exactFound.length > 0) {
      exactFound.sort((a, b) =>
        new Date(b.operationCompletedAt || 0) - new Date(a.operationCompletedAt || 0)
      );
      const exact = exactFound[0];

      result.productName = exact.product?.name || null;
      result.nomenclatureCode = exact.product?.nomenclatureCode || null;
      result.productBarcode = pickProductBarcode(exact, barcodeNorm);
      result.violator =
        fioFromUser(exact.responsibleUser) ||
        fioFromUser(exact.executor) ||
        null;
      result.violatorId =
        exact.responsibleUser?.id ||
        exact.executorId ||
        null;
      result.handlingUnitBarcode =
        exact?.targetAddress?.handlingUnitBarcode ||
        exact?.sourceAddress?.handlingUnitBarcode ||
        null;
      result.operationType = exact.operationType || null;
      result.operationCompletedAt = exact.operationCompletedAt || null;
      result.strategy = exactMatchMode === 'handling_unit_barcode'
        ? 'exact_cell_and_handling_unit_barcode'
        : 'exact_cell_and_barcode';
      return result;
    }
  } catch (e) {
    console.log('[Консолидация] Ошибка точного поиска по ячейке+ШК, продолжаем fallback');
  }

  console.log(`[Консолидация] Шаг A: ищем товар по ШК "${barcodeNorm}"...`);

  // ─── Поиск с пагинацией по месту, затем фильтрация по barcode на клиенте ───
  if (itemsA.length === 0) {
    console.log(`[Консолидация] Поиск по месту с пагинацией...`);

    const pageSize = 500;
    let pageNumber = 1;
    let found = [];
    let foundByHandlingUnit = false;
    const fallbackBodies = [];

    if (cellIds.length > 0) {
      for (const id of cellIds) {
        fallbackBodies.push({
          ...baseBody,
          targetCellId: id,
          operationTypes: LOOKUP_OPERATION_TYPES,
        });
      }
    } else {
      fallbackBodies.push({
        ...baseBody,
        operationTypes: LOOKUP_OPERATION_TYPES,
      });
    }

    while (true) {
      console.log(`[Консолидация] Загружаем страницу ${pageNumber}...`);

      const batches = await Promise.all(
        fallbackBodies.map(body => wmsPost(token, { ...body, pageNumber, pageSize }))
      );
      const allItems = batches.flatMap(b => b?.value?.items || []);
      if (allItems.length === 0) break;

      const byProductBarcode = allItems.filter(it => matchesBarcode(it, barcodeNorm));
      if (byProductBarcode.length > 0) {
        found = byProductBarcode;
        foundByHandlingUnit = false;
        console.log(`[Консолидация] Найдено на странице ${pageNumber}`);
        break;
      }

      const byHandlingUnit = allItems.filter(it => matchesHandlingUnitBarcode(it, barcodeNorm));
      if (byHandlingUnit.length > 0) {
        found = byHandlingUnit;
        foundByHandlingUnit = true;
        console.log(`[Консолидация] Найдено на странице ${pageNumber}`);
        break;
      }

      pageNumber++;
    }

    itemsA = found;
    strategy = found.length > 0
      ? (foundByHandlingUnit ? 'handling_unit_match_paginated' : 'ean_match_paginated')
      : 'not_found';
    if (itemsA.length === 0) console.log(`[Консолидация] Товар не найден во всех страницах`);
  }

  result.strategy = strategy;

  if (itemsA.length === 0) return result;

  // ─── Берём самую свежую операцию ─────────────
  itemsA.sort((a, b) =>
    new Date(b.operationCompletedAt || 0) - new Date(a.operationCompletedAt || 0)
  );

  const first = itemsA[0];

  result.productName = first.product?.name || null;
  result.nomenclatureCode = first.product?.nomenclatureCode || null;

  const productId = first.product?.productId ?? first.productId ?? null;
  console.log(`[Консолидация] productId: ${productId}`);

  if (!productId || !cell) return result;

  // ─── Шаг B — поиск нарушителя ─────────────
  console.log(`[Консолидация] Шаг B: ищем по ячейке "${cell}"`);

  const stepBQueries = [];
  if (cellIds.length > 0) {
    for (const id of cellIds) {
      stepBQueries.push(
        wmsPost(token, {
          ...baseBody,
          productId,
          targetCellId: id,
          pageNumber: 1,
          pageSize: 500,
        })
      );
    }
  } else {
    stepBQueries.push(
      wmsPost(token, {
        ...baseBody,
        productId,
        sourceCellId: null,
        targetCellId: null,
        pageNumber: 1,
        pageSize: 500,
      })
    );
  }
  const stepBData = await Promise.all(stepBQueries);
  const itemsB = stepBData.flatMap(x => x?.value?.items || []);

  const matched = itemsB.filter(it => {
    return matchesTargetCell(it, cellIds, cellNorm) && matchesBarcode(it, barcodeNorm);
  });
  const matchedByHandlingUnit = matched.length > 0
    ? matched
    : itemsB.filter(it => {
        return matchesTargetCell(it, cellIds, cellNorm) && matchesHandlingUnitBarcode(it, barcodeNorm);
      });

  if (matchedByHandlingUnit.length > 0) {
    matchedByHandlingUnit.sort((a, b) =>
      new Date(b.operationCompletedAt || 0) - new Date(a.operationCompletedAt || 0)
    );

    const v = matchedByHandlingUnit[0];

    result.violator =
      fioFromUser(v.responsibleUser) ||
      fioFromUser(v.executor) ||
      null;

    result.violatorId =
      v.responsibleUser?.id ||
      v.executorId ||
      null;
    result.productBarcode = pickProductBarcode(v, barcodeNorm);
    result.handlingUnitBarcode =
      v?.targetAddress?.handlingUnitBarcode ||
      v?.sourceAddress?.handlingUnitBarcode ||
      null;

    result.operationType = v.operationType || null;
    result.operationCompletedAt = v.operationCompletedAt || null;

    console.log(`[Консолидация] Нарушитель найден: ${result.violator}`);
  } else {
    console.log(`[Консолидация] По ячейке совпадений нет`);
  }

  return result;
}

// ─── Инициализация и загрузка ───────────────────────────────────────────────

export function initConsolidation() {
  const refreshBtn = document.getElementById('cons-refresh-btn');
  if (refreshBtn) refreshBtn.addEventListener('click', () => loadComplaints());

  const lookupAllBtn = document.getElementById('cons-lookup-all-btn');
  if (lookupAllBtn) lookupAllBtn.addEventListener('click', () => lookupAll());

  const sendTgBtn = document.getElementById('cons-send-telegram-btn');
  if (sendTgBtn) sendTgBtn.addEventListener('click', () => sendSelectedToTelegram(sendTgBtn));
  const bulkStatusBtn = document.getElementById('cons-bulk-apply-status');
  if (bulkStatusBtn) bulkStatusBtn.addEventListener('click', () => bulkApplyStatus(bulkStatusBtn));
  const bulkLookupBtn = document.getElementById('cons-bulk-lookup');
  if (bulkLookupBtn) bulkLookupBtn.addEventListener('click', () => bulkLookupSelected(bulkLookupBtn));
  const bulkDeleteBtn = document.getElementById('cons-bulk-delete');
  if (bulkDeleteBtn) bulkDeleteBtn.addEventListener('click', () => bulkDeleteSelected(bulkDeleteBtn));

  if (!modalControlsBound) {
    const prevBtn = document.getElementById('cons-photo-prev');
    const nextBtn = document.getElementById('cons-photo-next');
    prevBtn?.addEventListener('click', e => {
      e.stopPropagation();
      if (modalPhotoUrls.length <= 1) return;
      modalPhotoIndex = (modalPhotoIndex - 1 + modalPhotoUrls.length) % modalPhotoUrls.length;
      renderPhotoModalState();
    });
    nextBtn?.addEventListener('click', e => {
      e.stopPropagation();
      if (modalPhotoUrls.length <= 1) return;
      modalPhotoIndex = (modalPhotoIndex + 1) % modalPhotoUrls.length;
      renderPhotoModalState();
    });
    document.addEventListener('keydown', e => {
      const modal = document.getElementById('cons-photo-modal');
      if (!modal || !modal.classList.contains('modal--open')) return;
      if (e.key === 'ArrowLeft' && modalPhotoUrls.length > 1) {
        modalPhotoIndex = (modalPhotoIndex - 1 + modalPhotoUrls.length) % modalPhotoUrls.length;
        renderPhotoModalState();
      } else if (e.key === 'ArrowRight' && modalPhotoUrls.length > 1) {
        modalPhotoIndex = (modalPhotoIndex + 1) % modalPhotoUrls.length;
        renderPhotoModalState();
      } else if (e.key === 'Escape') {
        closePhotoModal();
      }
    });
    modalControlsBound = true;
  }

  const filterWrap = document.getElementById('cons-filters');
  if (filterWrap) {
    filterWrap.addEventListener('click', e => {
      const chip = e.target.closest('.cons-filter-chip');
      if (!chip) return;
      statusFilter = chip.dataset.filter;
      filterWrap.querySelectorAll('.cons-filter-chip').forEach(c =>
        c.classList.toggle('active', c.dataset.filter === statusFilter)
      );
      renderComplaints();
    });
  }
}

async function lookupAll() {
  const token = getToken();
  if (!token) { alert('Войдите в систему для поиска в WMS'); return; }

  const btn = document.getElementById('cons-lookup-all-btn');
  const needLookup = allComplaints.filter(c => !c.lookupDone);
  if (needLookup.length === 0) {
    // Повторить для всех — перезаписать результаты
    if (!confirm('Все жалобы уже проверены. Проверить заново?')) return;
    needLookup.push(...allComplaints);
  }

  btn.disabled = true;
  const total = needLookup.length;

  for (let i = 0; i < total; i++) {
    const c = needLookup[i];
    btn.textContent = `🔍 ${i + 1}/${total}...`;
    try {
      const result = await lookupViaBrowser(token, c.barcode, c.cell);
      await saveComplaintLookup(c.id, result);
    } catch (err) {
      await saveComplaintLookup(c.id, { lookupDone: false, lookupError: err.message || 'Ошибка WMS' });
    }
  }

  btn.disabled = false;
  btn.textContent = '🔍 Проверить все';
  await loadComplaints();
}

export async function loadComplaints() {
  try {
    allComplaints = await getConsolidationComplaints();
    const validIds = new Set(allComplaints.map(c => String(c.id)));
    selectedComplaintIds = new Set([...selectedComplaintIds].filter(id => validIds.has(id)));
    renderComplaints();
  } catch (err) {
    console.error('loadComplaints', err);
  }
}

async function sendSelectedToTelegram(btn) {
  const selected = allComplaints.filter(c => selectedComplaintIds.has(String(c.id)));
  if (selected.length === 0) {
    alert('Отметьте жалобы галочкой');
    return;
  }

  const inProgress = selected.filter(c => c.status === 'in_progress');
  if (inProgress.length === 0) {
    alert('Для отправки в Telegram выберите жалобы со статусом "В работе"');
    return;
  }

  const ids = inProgress.map(c => String(c.id));
  const prevText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Отправка...';
  try {
    const res = await sendComplaintsToTelegram(ids);
    if (!res?.ok) {
      const firstFailed = Array.isArray(res?.failed) && res.failed.length > 0
        ? res.failed[0]?.error
        : '';
      const msg = firstFailed || res?.error || 'Ошибка отправки в Telegram';
      alert(msg);
      return;
    }
    const sentCount = res.sentCount || 0;
    const failedCount = res.failedCount || 0;
    if (failedCount > 0) {
      const previewErrors = (Array.isArray(res.failed) ? res.failed : [])
        .slice(0, 3)
        .map(x => x?.error)
        .filter(Boolean)
        .join('\n');
      alert(`Отправлено: ${sentCount}, с ошибкой: ${failedCount}${previewErrors ? `\n\n${previewErrors}` : ''}`);
    } else {
      alert(`Отправлено в Telegram: ${sentCount}`);
    }
  } catch (err) {
    alert('Ошибка отправки: ' + (err?.message || err));
  } finally {
    btn.disabled = false;
    btn.textContent = prevText;
  }
}

function getSelectedComplaints() {
  return allComplaints.filter(c => selectedComplaintIds.has(String(c.id)));
}

async function bulkApplyStatus(btn) {
  const selected = getSelectedComplaints();
  if (selected.length === 0) {
    alert('Отметьте жалобы галочкой');
    return;
  }
  const statusEl = document.getElementById('cons-bulk-status');
  const status = statusEl ? statusEl.value : '';
  if (!['new', 'in_progress', 'resolved'].includes(status)) {
    alert('Выберите корректный статус');
    return;
  }
  const prevText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Сохранение...';
  try {
    const results = await Promise.allSettled(
      selected.map(c => updateComplaintStatus(c.id, status))
    );
    const ok = results.filter(r => r.status === 'fulfilled').length;
    const fail = results.length - ok;
    if (fail > 0) alert(`Статус обновлён: ${ok}, с ошибкой: ${fail}`);
    await loadComplaints();
  } catch (err) {
    alert('Ошибка: ' + (err?.message || err));
  } finally {
    btn.disabled = false;
    btn.textContent = prevText;
  }
}

async function bulkLookupSelected(btn) {
  const selected = getSelectedComplaints();
  if (selected.length === 0) {
    alert('Отметьте жалобы галочкой');
    return;
  }
  const token = getToken();
  if (!token) {
    alert('Войдите в систему для поиска в WMS');
    return;
  }
  const prevText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Проверка...';
  let ok = 0;
  let fail = 0;
  try {
    for (const c of selected) {
      try {
        const result = await lookupViaBrowser(token, c.barcode, c.cell);
        await saveComplaintLookup(c.id, result);
        ok++;
      } catch (err) {
        await saveComplaintLookup(c.id, {
          lookupDone: false,
          lookupError: err.message || 'Ошибка WMS',
        });
        fail++;
      }
    }
    if (fail > 0) alert(`Проверено: ${ok}, с ошибкой: ${fail}`);
    await loadComplaints();
  } finally {
    btn.disabled = false;
    btn.textContent = prevText;
  }
}

async function bulkDeleteSelected(btn) {
  const selected = getSelectedComplaints();
  if (selected.length === 0) {
    alert('Отметьте жалобы галочкой');
    return;
  }
  if (!confirm(`Удалить выбранные жалобы (${selected.length})?`)) return;
  const prevText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Удаление...';
  try {
    const results = await Promise.allSettled(
      selected.map(c => deleteComplaint(c.id))
    );
    const ok = results.filter(r => r.status === 'fulfilled').length;
    const fail = results.length - ok;
    if (fail > 0) alert(`Удалено: ${ok}, с ошибкой: ${fail}`);
    selected.forEach(c => selectedComplaintIds.delete(String(c.id)));
    await loadComplaints();
  } finally {
    btn.disabled = false;
    btn.textContent = prevText;
  }
}

function getFilteredComplaints() {
  if (statusFilter === 'all') return allComplaints;
  return allComplaints.filter(c => c.status === statusFilter);
}

// ─── Рендеринг ──────────────────────────────────────────────────────────────

function renderComplaints() {
  const container = document.getElementById('cons-table-wrap');
  if (!container) return;

  const filtered = getFilteredComplaints();

  const totalEl = document.getElementById('cons-count-total');
  const newEl = document.getElementById('cons-count-new');
  if (totalEl) totalEl.textContent = allComplaints.length;
  if (newEl) newEl.textContent = allComplaints.filter(c => c.status === 'new').length;
  const selectedEl = document.getElementById('cons-count-selected');
  if (selectedEl) selectedEl.textContent = selectedComplaintIds.size;

  if (filtered.length === 0) {
    container.innerHTML = '<div class="cons-empty">Нет жалоб</div>';
    return;
  }

  const rows = filtered.map(c => {
    const photos = Array.isArray(c.photoFilenames) && c.photoFilenames.length > 0
      ? c.photoFilenames
      : (c.photoFilename ? [c.photoFilename] : []);
    const photoCell = photos.length > 0
      ? (() => {
          const badge = photos.length > 1 ? `<span class="cons-photo-count-badge">${photos.length}</span>` : '';
          const imgs = photos.map((name, i) => {
            const url = `/api/consolidation/uploads/${esc(name)}`;
            const hidden = i > 0 ? ' cons-photo-thumb--hidden' : '';
            return `<img class="cons-photo-thumb${hidden}" src="${url}" alt="Фото" data-full="${url}">`;
          }).join('');
          return `<div class="cons-photo-stack">${badge}${imgs}</div>`;
        })()
      : '<span class="cons-no-photo">—</span>';

    const lookupInfo = c.lookupDone
      ? ''
      : c.lookupError
        ? `<span class="cons-lookup-err" title="${esc(c.lookupError)}">!</span>`
        : '';

    return `<tr data-id="${esc(c.id)}">
      <td class="cons-td-check">
        <input type="checkbox" class="cons-select-checkbox" data-id="${esc(c.id)}"${selectedComplaintIds.has(String(c.id)) ? ' checked' : ''}>
      </td>
      <td class="cons-td-date">${formatDate(c.createdAt)}</td>
      <td>${esc(c.employeeName || '—')}</td>
      <td class="cons-td-cell">${esc(c.cell)}</td>
      <td class="cons-td-barcode">${esc(c.barcode)}</td>
      <td>${esc(c.nomenclatureCode || '—')}</td>
      <td>${esc(c.productName || '—')}</td>
      <td>${esc(c.violator || '—')}${lookupInfo}</td>
      <td class="cons-td-date">${formatDate(c.operationCompletedAt)}</td>
      <td>${photoCell}</td>
      <td><span class="cons-status ${statusClass(c.status)}">${statusLabel(c.status)}</span></td>
      <td class="cons-actions">
        <select class="cons-status-select" data-id="${esc(c.id)}">
          <option value="new"${c.status === 'new' ? ' selected' : ''}>Новая</option>
          <option value="in_progress"${c.status === 'in_progress' ? ' selected' : ''}>В работе</option>
          <option value="resolved"${c.status === 'resolved' ? ' selected' : ''}>Решена</option>
        </select>
        <button class="btn btn-sm cons-btn-lookup" data-id="${esc(c.id)}" data-barcode="${esc(c.barcode)}" data-cell="${esc(c.cell)}" title="Поиск в WMS">&#x1f50d;</button>
        <button class="btn btn-sm cons-btn-delete" data-id="${esc(c.id)}" title="Удалить">&times;</button>
      </td>
    </tr>`;
  }).join('');

  container.innerHTML = `
    <table class="cons-table">
      <thead>
        <tr>
          <th><input type="checkbox" id="cons-select-all" title="Выбрать все"></th>
          <th>Дата</th>
          <th>Кто подал</th>
          <th>Место</th>
          <th>Штрихкод</th>
          <th>Артикул</th>
          <th>Товар</th>
          <th>Нарушитель</th>
          <th>Время нарушения</th>
          <th>Фото</th>
          <th>Статус</th>
          <th>Действия</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;

  attachHandlers(container);
}

// ─── Обработчики ─────────────────────────────────────────────────────────────

function attachHandlers(container) {
  const selectAll = container.querySelector('#cons-select-all');
  if (selectAll) {
    const rowChecks = [...container.querySelectorAll('.cons-select-checkbox')];
    const allChecked = rowChecks.length > 0 && rowChecks.every(cb => cb.checked);
    selectAll.checked = allChecked;
    selectAll.addEventListener('change', () => {
      rowChecks.forEach(cb => {
        cb.checked = selectAll.checked;
        const id = String(cb.dataset.id || '');
        if (!id) return;
        if (selectAll.checked) selectedComplaintIds.add(id);
        else selectedComplaintIds.delete(id);
      });
      const selectedEl = document.getElementById('cons-count-selected');
      if (selectedEl) selectedEl.textContent = selectedComplaintIds.size;
    });
  }

  container.querySelectorAll('.cons-select-checkbox').forEach(cb => {
    cb.addEventListener('change', () => {
      const id = String(cb.dataset.id || '');
      if (!id) return;
      if (cb.checked) selectedComplaintIds.add(id);
      else selectedComplaintIds.delete(id);
      const selectedEl = document.getElementById('cons-count-selected');
      if (selectedEl) selectedEl.textContent = selectedComplaintIds.size;
      const rowChecks = [...container.querySelectorAll('.cons-select-checkbox')];
      const selectAllEl = container.querySelector('#cons-select-all');
      if (selectAllEl) {
        selectAllEl.checked = rowChecks.length > 0 && rowChecks.every(x => x.checked);
      }
    });
  });

  // Status change
  container.querySelectorAll('.cons-status-select').forEach(sel => {
    sel.addEventListener('change', async () => {
      try {
        await updateComplaintStatus(sel.dataset.id, sel.value);
        await loadComplaints();
      } catch (err) { console.error('updateStatus', err); }
    });
  });

  // Delete
  container.querySelectorAll('.cons-btn-delete').forEach(btn => {
    btn.addEventListener('click', async () => {
      if (!confirm('Удалить жалобу?')) return;
      try {
        await deleteComplaint(btn.dataset.id);
        await loadComplaints();
      } catch (err) { console.error('delete', err); }
    });
  });

  // WMS lookup (из браузера)
  container.querySelectorAll('.cons-btn-lookup').forEach(btn => {
    btn.addEventListener('click', async () => {
      const token = getToken();
      if (!token) {
        alert('Войдите в систему для поиска в WMS');
        return;
      }
      btn.disabled = true;
      btn.textContent = '...';
      try {
        const result = await lookupViaBrowser(token, btn.dataset.barcode, btn.dataset.cell);
        await saveComplaintLookup(btn.dataset.id, result);
        await loadComplaints();
      } catch (err) {
        console.error('lookup', err);
        await saveComplaintLookup(btn.dataset.id, {
          lookupDone: false,
          lookupError: err.message || 'Ошибка WMS',
        });
        await loadComplaints();
      }
    });
  });

  // Photo modal
  container.querySelectorAll('.cons-photo-thumb').forEach(img => {
    img.addEventListener('click', (e) => {
      e.stopPropagation();
      const cell = img.closest('.cons-photo-stack') || img.parentElement;
      const thumbs = cell ? [...cell.querySelectorAll('.cons-photo-thumb')] : [img];
      const urls = thumbs.map(x => x.dataset.full).filter(Boolean);
      const idx = Math.max(0, urls.indexOf(img.dataset.full));
      openPhotoModal(urls, idx);
    });
  });
}

function renderPhotoModalState() {
  const modal = document.getElementById('cons-photo-modal');
  const modalImg = document.getElementById('cons-photo-modal-img');
  const counter = document.getElementById('cons-photo-counter');
  const prevBtn = document.getElementById('cons-photo-prev');
  const nextBtn = document.getElementById('cons-photo-next');
  if (!modal || !modalImg) return;
  if (!modalPhotoUrls.length) return;
  modalImg.src = modalPhotoUrls[modalPhotoIndex];
  if (counter) counter.textContent = `${modalPhotoIndex + 1} / ${modalPhotoUrls.length}`;
  const multi = modalPhotoUrls.length > 1;
  if (prevBtn) prevBtn.style.display = multi ? 'flex' : 'none';
  if (nextBtn) nextBtn.style.display = multi ? 'flex' : 'none';
}

function openPhotoModal(urls, index = 0) {
  const modal = document.getElementById('cons-photo-modal');
  if (!modal || !Array.isArray(urls) || urls.length === 0) return;
  modalPhotoUrls = urls;
  modalPhotoIndex = Math.min(Math.max(0, index), urls.length - 1);
  renderPhotoModalState();
  modal.classList.add('modal--open');
}

function closePhotoModal() {
  const modal = document.getElementById('cons-photo-modal');
  if (!modal) return;
  modal.classList.remove('modal--open');
}

// Close photo modal
document.addEventListener('click', e => {
  const modal = document.getElementById('cons-photo-modal');
  if (!modal) return;
  if (e.target === modal || e.target.classList.contains('cons-photo-modal-close')) {
    closePhotoModal();
  }
});
