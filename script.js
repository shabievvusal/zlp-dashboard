const API = '/api';
const TABLE_PAGE_SIZE = 100;
const IDLE_THRESHOLD_MS = 10 * 60 * 1000; // 10 минут

const el = (id) => document.getElementById(id);

let tableCurrentPage = 1;
let viewMode = 'stats'; // 'stats' | 'full'
let filterCompany = '__all__'; // '__all__' | '__none__' | company name
let shiftFilter = 'day'; // 'day' 9–21, 'night' 21–9
window._emplMap = null; // Map(normalizedFio -> company)
window._emplCompanies = []; // list of company names from CSV
window._lastTableData = []; // инициализируем сразу

// ==================== АВТОРИЗАЦИЯ ====================

let authToken = null;
let authTokenExpiry = null;
let refreshTokenValue = null;
let refreshTokenExpiry = null;
let refreshTimer = null;

async function login(login, password) {
  setStatus('Авторизация...', 'pending', 'Получаем токен');
  
  try {
    const response = await fetch('https://api.samokat.ru/wmsin-wwh/auth/password', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'Origin': 'https://wwh.samokat.ru',
        'Referer': 'https://wwh.samokat.ru/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Not(A:Brand";v="8", "Chromium";v="144", "Google Chrome";v="144"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"'
      },
      body: JSON.stringify({
        login,
        password
      })
    });

    if (!response.ok) {
      throw new Error(`Ошибка авторизации: ${response.status}`);
    }

    const data = await response.json();
    console.log('Auth response:', data);
    
    if (data.value && data.value.accessToken) {
      authToken = data.value.accessToken;
      authTokenExpiry = Date.now() + (data.value.expiresIn || 300) * 1000;
      
      if (data.value.refreshToken) {
        refreshTokenValue = data.value.refreshToken;
        refreshTokenExpiry = Date.now() + 30 * 24 * 60 * 60 * 1000;
      }
      
      await fetch(API + '/config', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          token: authToken,
          refreshToken: refreshTokenValue 
        }),
      });
      
      const tokenInput = el('token');
      if (tokenInput) tokenInput.value = authToken;
      
      const loginForm = el('login-form');
      const tokenForm = el('token-form');
      
      if (loginForm) loginForm.style.display = 'none';
      if (tokenForm) tokenForm.style.display = 'block';
      
      setStatus('Авторизация успешна', 'success', 'Токен получен (действует 5 мин)');
      
      scheduleTokenRefresh();
      
      return true;
    } else {
      console.error('Неожиданный ответ:', data);
      throw new Error('Токен не получен в ответе');
    }
  } catch (e) {
    setStatus('Ошибка авторизации', 'error', e.message);
    return false;
  }
}

async function refreshAccessToken() {
  if (!refreshTokenValue) {
    console.log('Нет refreshToken, требуется повторный вход');
    return false;
  }
  
  try {
    console.log('Обновление токена...');
    setStatus('Обновление токена...', 'pending', '');
    
    const response = await fetch('https://api.samokat.ru/wmsin-wwh/auth/refresh', {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'Origin': 'https://wwh.samokat.ru',
        'Referer': 'https://wwh.samokat.ru/',
      },
      body: JSON.stringify({
        refreshToken: refreshTokenValue
      })
    });

    if (!response.ok) {
      throw new Error(`Ошибка обновления токена: ${response.status}`);
    }

    const data = await response.json();
    
    if (data.value && data.value.accessToken) {
      authToken = data.value.accessToken;
      authTokenExpiry = Date.now() + (data.value.expiresIn || 300) * 1000;
      
      if (data.value.refreshToken) {
        refreshTokenValue = data.value.refreshToken;
        refreshTokenExpiry = Date.now() + 30 * 24 * 60 * 60 * 1000;
      }
      
      await fetch(API + '/config', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          token: authToken,
          refreshToken: refreshTokenValue 
        }),
      });
      
      const tokenInput = el('token');
      if (tokenInput) tokenInput.value = authToken;
      
      console.log('Токен успешно обновлен');
      setStatus('Токен обновлен', 'success', 'Действует еще 5 мин');
      
      return true;
    }
  } catch (e) {
    console.error('Ошибка обновления токена:', e);
    setStatus('Ошибка обновления токена', 'error', e.message);
  }
  
  return false;
}

function scheduleTokenRefresh() {
  if (refreshTimer) {
    clearTimeout(refreshTimer);
  }
  
  if (!authTokenExpiry) return;
  
  const timeUntilRefresh = Math.max(0, authTokenExpiry - Date.now() - 60 * 1000);
  
  refreshTimer = setTimeout(async () => {
    console.log('Плановое обновление токена...');
    const success = await refreshAccessToken();
    
    if (success) {
      scheduleTokenRefresh();
    } else {
      setStatus('Требуется авторизация', 'pending', 'Сессия истекла, войдите снова');
      
      const loginForm = el('login-form');
      const tokenForm = el('token-form');
      
      if (loginForm) loginForm.style.display = 'block';
      if (tokenForm) tokenForm.style.display = 'none';
      
      authToken = null;
      authTokenExpiry = null;
      refreshTokenValue = null;
      
      const tokenInput = el('token');
      if (tokenInput) tokenInput.value = '';
    }
  }, timeUntilRefresh);
  
  console.log(`Следующее обновление токена через ${Math.round(timeUntilRefresh / 1000)} сек`);
}

function logout() {
  authToken = null;
  authTokenExpiry = null;
  refreshTokenValue = null;
  refreshTokenExpiry = null;
  
  if (refreshTimer) {
    clearTimeout(refreshTimer);
    refreshTimer = null;
  }
  
  const tokenInput = el('token');
  if (tokenInput) tokenInput.value = '';
  
  const loginForm = el('login-form');
  const tokenForm = el('token-form');
  
  if (loginForm) loginForm.style.display = 'block';
  if (tokenForm) tokenForm.style.display = 'none';
  
  setStatus('Вы вышли из системы', 'pending', '');
}

function initAuthHandlers() {
  console.log('Инициализация обработчиков авторизации');
  
  const loginBtn = el('btn-login');
  if (loginBtn) {
    console.log('Кнопка входа найдена, добавляем обработчик');
    loginBtn.addEventListener('click', async function(e) {
      e.preventDefault();
      e.stopPropagation();
      
      console.log('Клик по кнопке входа');
      
      const loginInput = el('login');
      const passwordInput = el('password');
      
      if (!loginInput || !passwordInput) {
        console.error('Поля ввода не найдены');
        alert('Ошибка: поля ввода не найдены');
        return;
      }
      
      const loginValue = loginInput.value.trim();
      const passwordValue = passwordInput.value.trim();
      
      if (!loginValue || !passwordValue) {
        alert('Введите логин и пароль');
        return;
      }
      
      const success = await login(loginValue, passwordValue);
      console.log('Результат авторизации:', success);
    });
  } else {
    console.error('Кнопка входа не найдена!');
  }
  
  const logoutBtn = el('btn-logout');
  if (logoutBtn) {
    logoutBtn.addEventListener('click', function(e) {
      e.preventDefault();
      logout();
    });
  }
}

async function loadTokenFromConfig() {
  try {
    const r = await fetch(API + '/config');
    const data = await r.json();
    
    if (data.token && data.token !== '***') {
      authToken = data.token;
      authTokenExpiry = Date.now() + 60 * 1000;
      
      if (data.refreshToken) {
        refreshTokenValue = data.refreshToken;
        refreshTokenExpiry = Date.now() + 30 * 24 * 60 * 60 * 1000;
        
        const refreshed = await refreshAccessToken();
        if (refreshed) {
          const tokenInput = el('token');
          if (tokenInput) tokenInput.value = authToken;
          
          const loginForm = el('login-form');
          const tokenForm = el('token-form');
          
          if (loginForm) loginForm.style.display = 'none';
          if (tokenForm) tokenForm.style.display = 'block';
          
          console.log('Токен восстановлен из конфига');
          return;
        }
      }
    }
    
    const loginForm = el('login-form');
    const tokenForm = el('token-form');
    
    if (loginForm) loginForm.style.display = 'block';
    if (tokenForm) tokenForm.style.display = 'none';
    
  } catch (e) {
    console.error('Ошибка загрузки токена из конфига:', e);
  }
}

// ==================== РЕЖИМ ПОДКЛЮЧЕНИЯ ====================

let connectionMode = 'server'; // 'server' | 'browser'

function updateConnectionModeUI() {
  const serverBtn = el('btn-mode-server');
  const browserBtn = el('btn-mode-browser');
  const hint = el('connection-hint');
  
  if (serverBtn) serverBtn.classList.toggle('btn-mode-connection-active', connectionMode === 'server');
  if (browserBtn) browserBtn.classList.toggle('btn-mode-connection-active', connectionMode === 'browser');
  
  if (hint) {
    if (connectionMode === 'server') {
      hint.innerHTML = '✓ Дом/Офис - работает всегда, через сервер';
      hint.style.color = '#4CAF50';
    } else {
      hint.innerHTML = '⚠ VPN - требуется подключение к VPN Самокат';
      hint.style.color = '#ff9800';
    }
  }
  
  fetch(API + '/config', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ connectionMode }),
  }).catch(e => console.error('Ошибка сохранения режима:', e));
}

async function setConnectionMode(mode) {
  connectionMode = mode;
  updateConnectionModeUI();
  setStatus('Режим изменен', 'success', mode === 'server' ? 'Запросы через сервер' : 'Прямые запросы (нужен VPN)');
}

async function loadConnectionMode() {
  try {
    const r = await fetch(API + '/config');
    const data = await r.json();
    if (data.connectionMode) {
      connectionMode = data.connectionMode;
      updateConnectionModeUI();
    }
  } catch (e) {
    console.error('Ошибка загрузки режима:', e);
  }
}

// ==================== ОСНОВНЫЕ ФУНКЦИИ ====================

function isDayShift(operationCompletedAt) {
  if (!operationCompletedAt) return false;
  const h = new Date(operationCompletedAt).getHours();
  return h >= 9 && h < 21;
}

function isNightShift(operationCompletedAt) {
  if (!operationCompletedAt) return false;
  const h = new Date(operationCompletedAt).getHours();
  return h >= 21 || h < 9;
}

function setStatus(text, type = '', detail = '') {
  const status = el('status');
  const statusDetail = el('status-detail');
  if (status) {
    status.textContent = text;
    status.className = 'status' + (type ? ' ' + type : '');
  }
  if (statusDetail) statusDetail.textContent = detail;
}

function getToken() {
  const tokenEl = el('token');
  return tokenEl ? tokenEl.value.trim() : '';
}

function getCookie() {
  const cookieEl = el('cookie');
  return (cookieEl && cookieEl.value) ? cookieEl.value.trim() : '';
}

function getInterval() {
  const intervalEl = el('interval');
  return intervalEl ? (parseInt(intervalEl.value, 10) || 60) : 60;
}

function getDateRange() {
  const fromEl = el('date-from');
  const toEl = el('date-to');
  if (!fromEl || !toEl) return {};
  
  const from = fromEl.value;
  const to = toEl.value;
  if (!from || !to) return {};
  
  return {
    operationCompletedAtFrom: new Date(from).toISOString(),
    operationCompletedAtTo: new Date(to).toISOString(),
  };
}

function formatDateTimeForDisplay(isoStr) {
  if (!isoStr) return '—';
  const d = new Date(isoStr);
  const pad = n => String(n).padStart(2, '0');
  return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function updateDateDisplays() {
  const from = el('date-from')?.value || '';
  const to = el('date-to')?.value || '';
  const fromDisplay = document.getElementById('date-from-display');
  const toDisplay = document.getElementById('date-to-display');
  if (fromDisplay) fromDisplay.textContent = formatDateTimeForDisplay(from);
  if (toDisplay) toDisplay.textContent = formatDateTimeForDisplay(to);
}

function updateDateDisplaysFromPicker() {
  const { rangeFrom, rangeTo } = datePickerState;
  const timeFrom = parseTime(el('time-from')?.value) || '00:00';
  const timeTo = parseTime(el('time-to')?.value) || '23:59';
  const pad = n => String(n).padStart(2, '0');
  const fromDate = rangeFrom || new Date();
  const toDate = rangeTo || rangeFrom || new Date();
  const fromValue = `${fromDate.getFullYear()}-${pad(fromDate.getMonth() + 1)}-${pad(fromDate.getDate())}T${timeFrom}:00`;
  const toValue = `${toDate.getFullYear()}-${pad(toDate.getMonth() + 1)}-${pad(toDate.getDate())}T${timeTo}:00`;
  const fromInput = el('date-from');
  const toInput = el('date-to');
  if (fromInput) fromInput.value = fromValue;
  if (toInput) toInput.value = toValue;
  const fromDisplay = document.getElementById('date-from-display');
  const toDisplay = document.getElementById('date-to-display');
  if (fromDisplay) fromDisplay.textContent = formatDateTimeForDisplay(fromValue);
  if (toDisplay) toDisplay.textContent = formatDateTimeForDisplay(toValue);
}

function getFetchOptions() {
  const options = getDateRange();
  const modeEl = document.querySelector('input[name="fetch-mode"]:checked');
  const mode = modeEl?.value || 'single';
  
  if (mode === 'all') {
    options.fetchAllPages = true;
    options.pageSize = 1000;
  } else if (mode === 'limit') {
    const n = parseInt(el('max-rows')?.value, 10);
    if (n > 0) {
      options.fetchAllPages = true;
      options.maxRows = n;
      options.pageSize = 1000;
    }
  }
  
  if (!options.pageSize) options.pageSize = 100;
  return options;
}

const SAMOKAT_API_URL = 'https://api.samokat.ru/wmsops-wwh/stocks/changes/search';

function buildBodyForApi(options) {
  const from = options.operationCompletedAtFrom || new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString().slice(0, 10) + 'T21:00:00.000Z';
  const to = options.operationCompletedAtTo || new Date().toISOString().slice(0, 10) + 'T20:59:59.000Z';
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
    pageSize: options.pageSize || 100,
  };
}

function fetchViaBrowserChecked() {
  const cb = el('fetch-via-browser');
  return cb ? cb.checked : false;
}

function formatDate(d) {
  if (!d) return '—';
  const x = new Date(d);
  return x.toLocaleString('ru-RU');
}

let useVpn = true;

function updateSubtitle() {
  const sub = el('subtitle');
  if (sub) sub.textContent = useVpn ? 'Опрос API stocks/changes через VPN' : 'Опрос API stocks/changes (офис, без VPN)';
}

function updateVpnToggleUI() {
  const officeBtn = el('btn-mode-office');
  const vpnBtn = el('btn-mode-vpn');
  if (officeBtn) officeBtn.classList.toggle('btn-vpn-mode-active', !useVpn);
  if (vpnBtn) vpnBtn.classList.toggle('btn-vpn-mode-active', useVpn);
  updateSubtitle();
}

async function setVpnMode(vpn) {
  useVpn = !!vpn;
  try {
    await fetch(API + '/config', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ useVpn }),
    });
    updateVpnToggleUI();
  } catch (e) {
    setStatus('Ошибка сохранения режима', 'error', e.message);
  }
}

async function fetchStatus() {
  try {
    const r = await fetch(API + '/status');
    const data = await r.json();
    
    const intervalEl = el('interval');
    if (intervalEl) intervalEl.value = data.config?.intervalMinutes ?? 60;
    
    const cookieEl = el('cookie');
    if (cookieEl) cookieEl.value = data.config?.cookie || '';
    
    if (data.config?.useVpn !== undefined) {
      useVpn = !!data.config.useVpn;
      updateVpnToggleUI();
    }
    
    if (data.scheduleRunning) {
      setStatus('Статус: Автоопрос включён', 'success', data.lastRun ? 'Последний запуск: ' + formatDate(data.lastRun) : '');
    } else {
      setStatus('Статус: Остановлено', 'pending', '');
    }
    return data;
  } catch (e) {
    setStatus('Статус: Ошибка связи с сервером', 'error', e.message);
    return null;
  }
}

async function fetchOnePageFromBrowser(token, headers, baseBody, pageNum) {
  const body = { ...baseBody, pageNumber: pageNum };
  const r = await fetch(SAMOKAT_API_URL, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });
  
  const text = await r.text();
  const trimmed = text && text.trim().toLowerCase().replace(/\s+/g, ' ');
  
  if (trimmed.startsWith('<!doctype') || trimmed.startsWith('<html')) {
    throw new Error('Сервер вернул HTML вместо JSON');
  }
  
  let data;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    throw new Error('Ответ не JSON');
  }
  
  if (!r.ok) {
    const msg = data?.message || data?.error || r.statusText || text.slice(0, 150);
    throw new Error(`${r.status}: ${msg}`);
  }
  
  const value = data?.value || data;
  const items = Array.isArray(value?.items) ? value.items : (Array.isArray(data?.content) ? data.content : []);
  const total = value?.total ?? data?.totalElements ?? null;
  return { items, total };
}

async function fetchFromBrowser() {
  if (authTokenExpiry && authTokenExpiry < Date.now()) {
    console.log('Токен истек, пробуем обновить...');
    const refreshed = await refreshAccessToken();
    
    if (!refreshed) {
      setStatus('Ошибка', 'error', 'Токен истек, выполните повторный вход');
      
      const loginForm = el('login-form');
      const tokenForm = el('token-form');
      
      if (loginForm) loginForm.style.display = 'block';
      if (tokenForm) tokenForm.style.display = 'none';
      
      return null;
    }
  }
  
  const token = getToken();
  if (!token) {
    setStatus('Ошибка', 'error', 'Укажите Bearer токен в настройках.');
    return null;
  }
  
  const options = getFetchOptions();
  const pageSize = Math.min(1000, Math.max(100, parseInt(options.pageSize, 10) || 100));
  const maxRows = options.maxRows != null ? Math.max(1, parseInt(options.maxRows, 10) || 0) : null;
  const fetchAll = options.fetchAllPages || (maxRows != null && maxRows > 0);

  const baseBody = buildBodyForApi({
    ...getDateRange(),
    pageNumber: 1,
    pageSize,
  });
  
  const headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Origin': 'https://wwh.samokat.ru',
    'Referer': 'https://wwh.samokat.ru/',
    'Authorization': `Bearer ${token}`,
  };

  const startedAt = Date.now();
  let allItems = [];
  let totalFromApi = null;
  let pageNum = 1;

  try {
    const first = await fetchOnePageFromBrowser(token, headers, baseBody, 1);
    allItems = first.items;
    totalFromApi = first.total;

    if (fetchAll && totalFromApi != null && totalFromApi > allItems.length) {
      const totalPages = Math.ceil(totalFromApi / pageSize);
      
      const pagesToLoad = [];
      for (let p = 2; p <= totalPages; p++) {
        if (maxRows != null && allItems.length >= maxRows) {
          break;
        }
        pagesToLoad.push(p);
      }
      
      setStatus(`Загрузка... 1 из ${totalPages}`, 'pending', 
        `Всего записей: ${totalFromApi}, загружаем ${pagesToLoad.length} страниц`);
      
      const CONCURRENT = 5;
      for (let i = 0; i < pagesToLoad.length; i += CONCURRENT) {
        const batch = pagesToLoad.slice(i, i + CONCURRENT);
        
        const promises = batch.map(page => 
          fetchOnePageFromBrowser(token, headers, baseBody, page)
            .catch(err => {
              console.error(`Ошибка загрузки страницы ${page}:`, err);
              return { items: [] };
            })
        );
        
        const results = await Promise.all(promises);
        
        for (const result of results) {
          allItems = allItems.concat(result.items);
        }
        
        setStatus(`Загрузка... страницы ${i+1}-${Math.min(i+CONCURRENT, pagesToLoad.length)} из ${totalPages}`, 'pending', 
          `Загружено: ${allItems.length} / ${totalFromApi} записей`);
        
        if (maxRows != null && allItems.length >= maxRows) {
          allItems = allItems.slice(0, maxRows);
          break;
        }
      }
      
      pageNum = pagesToLoad.length + 1;
    }
  } catch (e) {
    setStatus('Ошибка', 'error', e.message);
    return null;
  }
  
  const total = totalFromApi ?? allItems.length;
  const resultItems = maxRows != null ? allItems.slice(0, maxRows) : allItems;
  const payload = {
    value: { items: resultItems, total },
    operationCompletedAtFrom: baseBody.operationCompletedAtFrom,
    operationCompletedAtTo: baseBody.operationCompletedAtTo,
  };
  let savedTo = '';
  
  try {
    const saveRes = await fetch(API + '/save-fetched-data', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const saveData = await saveRes.json();
    savedTo = saveData.savedTo || '';
  } catch (_) {}
  
  const result = {
    success: true,
    data: { value: { items: resultItems, total }, raw: { value: { items: resultItems, total } } },
    count: resultItems.length,
    total,
    duration: Date.now() - startedAt,
    savedTo,
    pagesFetched: pageNum,
  };
  return result;
}

async function fetchDataNow() {
  const btn = el('btn-fetch');
  if (btn) btn.disabled = true;
  setStatus('Запрос...', 'pending', '');

  const token = getToken();
  const options = getFetchOptions();

  try {
    if (connectionMode === 'browser' && fetchViaBrowserChecked()) {
      const result = await fetchFromBrowser();
      if (result) {
        setStatus(
          `Успех: ${result.count ?? 0} записей из ${result.total ?? 0}`,
          'success',
          result.savedTo ? `Сохранено в ${result.savedTo}` : 'Через браузер'
        );
        viewMode = 'stats';
        renderTable(result.data);
        addHistoryItem(result);
        if ('Notification' in window && Notification.permission === 'granted') {
          new Notification('OWMS', { body: `Получено записей: ${result.count ?? 0}` });
        }
      }
      if (btn) btn.disabled = false;
      return;
    }
  } catch (e) {
    setStatus('Ошибка', 'error', e.message);
    renderTable(null);
    if (btn) btn.disabled = false;
    return;
  }

  try {
    const body = { options };
    if (token) body.token = token;

    const r = await fetch(API + '/fetch-data', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const result = await r.json();

    if (result.success) {
      const total = result.data?.value?.total ?? result.total ?? result.count ?? 0;
      const pagesInfo = result.pagesFetched ? `, страниц: ${result.pagesFetched}` : '';
      setStatus(
        `Успех: ${result.count ?? 0} записей из ${total}`,
        'success',
        `Время: ${result.duration} мс${pagesInfo}, сохранено в ${result.savedTo || '—'}`
      );
      viewMode = 'stats';
      renderTable(result.data);
      addHistoryItem(result);
      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('OWMS', { body: `Получено записей: ${result.count ?? 0}` });
      }
    } else {
      const errMsg = result.error || r.statusText;
      const isHtmlError = errMsg && errMsg.includes('HTML');
      const detail = isHtmlError
        ? errMsg + ' Попробуйте переключиться на "Дом/Офис" режим.'
        : errMsg;
      setStatus('Ошибка запроса', 'error', detail);
      renderTable(null);
      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('OWMS — Ошибка', { body: result.error });
      }
    }
  } catch (e) {
    setStatus('Ошибка', 'error', e.message);
    renderTable(null);
  } finally {
    if (btn) btn.disabled = false;
  }
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
    barcodes: (p.barcodes || []).join(', ') || '—',
    productionDate: part.productionDate || '—',
    bestBeforeDate: part.bestBeforeDate || '—',
    cellAddress: src.cellAddress || item.cellAddress || '—',
    sourceBarcode: src.handlingUnitBarcode || '—',
    targetCellAddress: tgt.cellAddress || '—',
    targetBarcode: tgt.handlingUnitBarcode || '—',
    operationStartedAt: item.operationStartedAt || null,
    operationCompletedAt: item.operationCompletedAt || null,
    responsibleUser: [ru.lastName, ru.firstName, ru.middleName].filter(Boolean).join(' ') || '—',
    sourceOld: (item.sourceQuantity || {}).oldQuantity || 0,
    sourceNew: (item.sourceQuantity || {}).newQuantity || 0,
    targetOld: (item.targetQuantity || {}).oldQuantity || 0,
    targetNew: (item.targetQuantity || {}).newQuantity || 0,
  };
}

const TABLE_HEADERS = {
  productName: 'Товар',
  nomenclatureCode: 'Код номенклатуры',
  productId: 'ID товара',
  barcodes: 'Штрихкоды',
  productionDate: 'Дата пр-ва',
  bestBeforeDate: 'Годен до',
  cellAddress: 'Ячейка',
  sourceBarcode: 'ШК источника',
  targetBarcode: 'ШК приёмника',
  operationStartedAt: 'Начало операции',
  operationCompletedAt: 'Окончание операции',
  responsibleUser: 'Ответственный',
  sourceOld: 'Исх. было',
  sourceNew: 'Исх. стало',
  targetOld: 'Приём было',
  targetNew: 'Приём стало',
  id: 'ID',
  type: 'Тип',
  operationType: 'Тип операции',
};

function formatTimeOnly(isoStr) {
  if (!isoStr) return '—';
  const d = new Date(isoStr);
  return d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function normalizeFio(s) {
  return (s || '').trim().replace(/\s+/g, ' ');
}

function fioToKey(fio) {
  const parts = normalizeFio(fio).split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).join(' ');
}

async function loadEmpl() {
  try {
    const r = await fetch(API + '/empl');
    const data = await r.json();
    window._emplMap = new Map();
    (data.employees || []).forEach(({ fio, company }) => {
      const key = fioToKey(fio);
      if (key) window._emplMap.set(key, (company || '').trim());
    });
    window._emplCompanies = data.companies || [];
    return true;
  } catch {
    window._emplMap = new Map();
    window._emplCompanies = [];
    return false;
  }
}

function getCompanyForFio(fio) {
  if (!window._emplMap) return '';
  const key = fioToKey(fio);
  return window._emplMap.get(key) ?? '';
}

// ==================== ИСПРАВЛЕННАЯ СТАТИСТИКА С ФИЛЬТРАЦИЕЙ ПО ДАТАМ ====================

function computeStats(content) {
  if (!content || !Array.isArray(content) || content.length === 0) {
    return [];
  }
  
  // 🔥 ПОЛУЧАЕМ ДИАПАЗОН ДАТ ИЗ ФИЛЬТРА
  const dateRange = getDateRange();
  let fromDate = 0;
  let toDate = Infinity;
  
  if (dateRange.operationCompletedAtFrom && dateRange.operationCompletedAtTo) {
    fromDate = new Date(dateRange.operationCompletedAtFrom).getTime();
    toDate = new Date(dateRange.operationCompletedAtTo).getTime();
    console.log('Фильтр дат:', new Date(fromDate).toISOString(), '-', new Date(toDate).toISOString());
  }
  
  // 🔥 ФИЛЬТРУЕМ КОНТЕНТ ПО ДАТЕ ЗАВЕРШЕНИЯ ОПЕРАЦИИ
  const filteredContent = content.filter(row => {
    if (!row.operationCompletedAt) return false;
    const rowDate = new Date(row.operationCompletedAt).getTime();
    return rowDate >= fromDate && rowDate <= toDate;
  });
  
  console.log(`Исходных записей: ${content.length}, после фильтра по дате: ${filteredContent.length}`);
  
  if (!filteredContent || !Array.isArray(filteredContent) || filteredContent.length === 0) {
    return [];
  }
  
  const byUser = new Map();
  
  for (const row of filteredContent) {
    let fio = row.responsibleUser || '';
    fio = fio.trim();
    if (!fio || fio === '—' || fio === '') {
      fio = 'Неизвестно';
    }
    
    if (!byUser.has(fio)) byUser.set(fio, []);
    byUser.get(fio).push(row);
  }
  
  const stats = [];
  
  for (const [fio, rows] of byUser) {
    const company = getCompanyForFio(fio) || '—';
    
    // ХР - считаем ВСЕ строки (никаких дублей!)
    let pieceSelectionCount = 0;
    
    // 🔥 КДК - удаляем дубли по ТОВАР + ЯЧЕЙКА НАЗНАЧЕНИЯ
    const uniqueKdkTasksSet = new Set();
    
    let pieces = 0;
    const eoSet = new Set();
    
    for (const r of rows) {
      // ХР - просто счетчик
      if (r.operationType === 'PIECE_SELECTION_PICKING') {
        pieceSelectionCount++;
      }
      
      // 🔥 КДК - уникальные по товару + целевой ячейке
      if (r.operationType === 'PICK_BY_LINE') {
        const productId = r.productId || 'no-product';
        const targetCell = r.targetCellAddress || 'no-target-cell';
        const kdkTaskKey = `${productId}||${targetCell}`;
        uniqueKdkTasksSet.add(kdkTaskKey);
      }
      
      // Суммируем штуки
      const targetNew = Number(r.targetNew) || 0;
      pieces += targetNew;
      
      // Уникальные ЕО (просто для статистики)
      if (r.targetBarcode && r.targetBarcode.trim() && r.targetBarcode !== '—') {
        eoSet.add(r.targetBarcode.trim());
      }
    }
    
    const kdk = uniqueKdkTasksSet.size;        // ✅ уникальные товар+целевая ячейка
    const hr = pieceSelectionCount;             // ✅ все строки ХР (без дублей!)
    const sz = pieceSelectionCount + kdk;       // ✅ СЗ = все ХР + уникальные КДК
    const eo = eoSet.size;
    
    const sorted = [...rows].sort((a, b) => {
      const aTime = a.operationCompletedAt ? new Date(a.operationCompletedAt).getTime() : 0;
      const bTime = b.operationCompletedAt ? new Date(b.operationCompletedAt).getTime() : 0;
      return aTime - bTime;
    });
    
    const idles = [];
    for (let i = 1; i < sorted.length; i++) {
      const prevEnd = sorted[i - 1].operationCompletedAt ? 
        new Date(sorted[i - 1].operationCompletedAt).getTime() : 0;
      const nextStart = sorted[i].operationCompletedAt ? 
        new Date(sorted[i].operationCompletedAt).getTime() : 0;
      
      if (prevEnd && nextStart && (nextStart - prevEnd >= IDLE_THRESHOLD_MS)) {
        idles.push(
          formatTimeOnly(sorted[i - 1].operationCompletedAt) + 
          '–' + 
          formatTimeOnly(sorted[i].operationCompletedAt)
        );
      }
    }
    const idlesStr = idles.length ? idles.join(', ') : '—';
    
    stats.push({
      fio,
      company,
      sz,
      hr,
      kdk,
      pieces,
      idlesStr,
      eo
    });
  }
  
  stats.sort((a, b) => (b.sz - a.sz) || (b.pieces - a.pieces));
  return stats;
}

/** СЗ по часам: для каждого (ФИО, час) считаем СЗ так же, как в computeStats (ХР + уникальные КДК). */
const DSH_HOURS = Array.from({ length: 12 }, (_, i) => 10 + i); // 10..21
function computeSzByHour(content) {
  if (!content || !Array.isArray(content) || content.length === 0) {
    return { fios: [], hours: DSH_HOURS, getSz: () => 0 };
  }
  // Столбец 10 = интервал 9:00–10:00, столбец 11 = 10:00–11:00 и т.д. (ключ = час окончания интервала)
  const byFio = new Map();
  for (const row of content) {
    if (!row.operationCompletedAt) continue;
    const dt = new Date(row.operationCompletedAt);
    const hour = dt.getHours();           // 0..23
    const col = hour + 1;                 // 9→10, 10→11, ..., 20→21
    if (col < 10 || col > 21) continue;
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
  const getTotal = (f) => DSH_HOURS.reduce((s, h) => s + (byFio.get(f).get(h)?.pieceSelectionCount || 0) + (byFio.get(f).get(h)?.kdkSet?.size || 0), 0);
  const fios = [...byFio.keys()].sort((a, b) => {
    const companyA = getCompanyForFio(a) || '—';
    const companyB = getCompanyForFio(b) || '—';
    const byCompany = (companyA).localeCompare(companyB);
    if (byCompany !== 0) return byCompany;
    return getTotal(b) - getTotal(a);
  });
  function getSz(fio, hour) {
    const byHour = byFio.get(fio);
    if (!byHour) return 0;
    const cell = byHour.get(hour);
    if (!cell) return 0;
    return cell.pieceSelectionCount + (cell.kdkSet ? cell.kdkSet.size : 0);
  }
  return { fios, hours: DSH_HOURS, getSz };
}

/** Данные для дашборда: текущие с страницы или загруженные из файла (уже flatten). */
let _dshContent = null;

function getDashboardContent() {
  return _dshContent || window._lastTableData || [];
}

function setDashboardContent(items) {
  _dshContent = Array.isArray(items) ? items : [];
}

/** Класс ячейки по значению СЗ в статистике: >=750 белый, 350–749 от жёлтого к красному, <350 красный */
function szCellClass(v) {
  const n = Number(v) || 0;
  if (n >= 750) return 'sz-white';
  if (n < 350) return 'sz-red';
  if (n < 550) return 'sz-orange';  /* 350–549 */
  return 'sz-yellow';               /* 550–749 */
}

/** Подсветка для «СЗ по часам»: без изменений — <50 красный, 51–74 жёлтый, 75+ белый */
function szCellClassHours(v) {
  const n = Number(v) || 0;
  if (n < 50) return 'sz-red';
  if (n <= 74) return 'sz-yellow';
  return 'sz-white';
}

let dshViewMode = 'stats'; // 'stats' | 'hours'

function renderDashboard() {
  const content = getDashboardContent();
  const thead = el('dsh-thead');
  const tbody = el('dsh-tbody');
  const emptyEl = el('dsh-empty');
  const sourceInfo = el('dsh-source-info');
  if (!thead || !tbody) return;

  if (sourceInfo) sourceInfo.textContent = content.length ? `Записей: ${content.length}` : '';

  if (!content.length) {
    thead.innerHTML = '';
    tbody.innerHTML = '';
    if (emptyEl) emptyEl.style.display = 'block';
    return;
  }
  if (emptyEl) emptyEl.style.display = 'none';

  const stats = computeStats(content);
  stats.sort((a, b) => {
    const companyA = (a.company || '—').localeCompare(b.company || '—');
    if (companyA !== 0) return companyA;
    return (b.sz || 0) - (a.sz || 0);
  });

  const { fios, hours, getSz } = computeSzByHour(content);

  if (dshViewMode === 'stats') {
    thead.innerHTML = '<tr><th>Компания</th><th>ФИО</th><th>СЗ</th><th>ХР</th><th>КДК</th><th>ШТ</th><th>Простои &gt;10 мин</th><th>ЕО</th></tr>';
    tbody.innerHTML = stats.map(row => {
      const sz = row.sz || 0;
      const szClass = szCellClass(sz);
      return `<tr>
        <td>${escapeHtml(row.company || '—')}</td>
        <td><strong>${escapeHtml(row.fio)}</strong></td>
        <td class="number-cell sz-cell ${szClass}">${sz}</td>
        <td class="number-cell">${row.hr || 0}</td>
        <td class="number-cell">${row.kdk || 0}</td>
        <td class="number-cell">${row.pieces || 0}</td>
        <td>${escapeHtml(row.idlesStr || '—')}</td>
        <td class="number-cell">${row.eo || 0}</td>
      </tr>`;
    }).join('');
  } else {
    thead.innerHTML = '<tr><th>Компания</th><th>ФИО</th>' + hours.map(h => `<th>${h}</th>`).join('') + '<th>Σ</th></tr>';
    tbody.innerHTML = fios.map(fio => {
      const company = getCompanyForFio(fio) || '—';
      const rowVals = hours.map(h => getSz(fio, h));
      const total = rowVals.reduce((a, b) => a + b, 0);
      const totalClass = szCellClassHours(total);
      return '<tr><td>' + escapeHtml(company) + '</td><td><strong>' + escapeHtml(fio) + '</strong></td>' +
        rowVals.map(v => `<td class="number-cell sz-cell ${szCellClassHours(v)}">${v}</td>`).join('') +
        `<td class="number-cell sz-cell ${totalClass}"><strong>${total}</strong></td></tr>`;
    }).join('');
  }
}

function showPage(page) {
  const main = document.getElementById('page-main');
  const dsh = document.getElementById('page-dsh');
  const se = document.getElementById('page-se');
  const navMain = el('nav-main');
  const navDsh = el('nav-dsh');
  const navSe = el('nav-se');
  if (page === 'dsh') {
    if (main) main.style.display = 'none';
    if (dsh) dsh.style.display = 'block';
    if (se) se.style.display = 'none';
    if (navMain) navMain.classList.remove('nav-link-active');
    if (navDsh) navDsh.classList.add('nav-link-active');
    if (navSe) navSe.classList.remove('nav-link-active');
    fillDshHistorySelect();
    document.querySelectorAll('.dsh-toggle-btn').forEach(b => {
      b.classList.toggle('dsh-toggle-active', (b.getAttribute('data-view') === dshViewMode));
    });
    renderDashboard();
  } else if (page === 'se') {
    if (main) main.style.display = 'none';
    if (dsh) dsh.style.display = 'none';
    if (se) se.style.display = 'block';
    if (navMain) navMain.classList.remove('nav-link-active');
    if (navDsh) navDsh.classList.remove('nav-link-active');
    if (navSe) navSe.classList.add('nav-link-active');
    renderSettingsPage();
  } else {
    if (main) main.style.display = 'block';
    if (dsh) dsh.style.display = 'none';
    if (se) se.style.display = 'none';
    if (navMain) navMain.classList.add('nav-link-active');
    if (navDsh) navDsh.classList.remove('nav-link-active');
    if (navSe) navSe.classList.remove('nav-link-active');
  }
}

function getCurrentPage() {
  const path = (location.pathname || '').trim();
  const hash = (location.hash || '').trim().toLowerCase();
  if (path === '/dsh' || hash === '#/dsh') return 'dsh';
  if (path === '/se' || hash === '#/se') return 'se';
  return 'main';
}

function renderSettingsPage() {
  const listEl = el('empl-no-company-list');
  const emptyEl = el('empl-no-company-empty');
  if (!listEl) return;
  const content = getDashboardContent();
  const withCompany = new Set();
  if (window._emplMap) {
    for (const [key] of window._emplMap) withCompany.add(key);
  }
  const fioToFull = new Map();
  for (const row of content) {
    const fio = (row.responsibleUser || '').trim();
    if (!fio || fio === '—') continue;
    const key = fioToKey(fio);
    if (!key) continue;
    if (!withCompany.has(key)) fioToFull.set(key, fio);
  }
  const noCompany = [...fioToFull.entries()].map(([k, fio]) => fio).sort((a, b) => a.localeCompare(b));
  if (noCompany.length === 0) {
    listEl.innerHTML = '';
    if (emptyEl) emptyEl.style.display = 'block';
    return;
  }
  if (emptyEl) emptyEl.style.display = 'none';
  listEl.innerHTML = noCompany.map(f => `<li><button type="button" class="btn-empl-fio">${escapeHtml(f)}</button></li>`).join('');
  listEl.querySelectorAll('.btn-empl-fio').forEach(btn => {
    btn.addEventListener('click', async () => {
      const fio = btn.textContent.trim();
      const company = window.prompt('Введите компанию для сотрудника:\n' + fio, '');
      if (company == null) return;
      try {
        const r = await fetch(API + '/empl', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fio, company: company.trim() }),
        });
        const contentType = (r.headers.get('Content-Type') || '').toLowerCase();
        if (!contentType.includes('application/json')) {
          const text = await r.text();
          console.error('POST /api/empl не вернул JSON:', text.slice(0, 200));
          alert('Сервер вернул не JSON. Убедитесь, что приложение запущено через node (npm start или start.sh), а не через другой сервер. Перезапустите backend и обновите страницу.');
          return;
        }
        const data = await r.json();
        if (data.ok) {
          await loadEmpl();
          renderSettingsPage();
        } else {
          alert('Ошибка: ' + (data.error || 'не удалось сохранить'));
        }
      } catch (e) {
        alert('Ошибка: ' + e.message);
      }
    });
  });
}

async function fillDshHistorySelect() {
  const sel = el('dsh-history-file');
  if (!sel) return;
  try {
    const r = await fetch(API + '/history');
    const files = await r.json();
    const currentVal = sel.value;
    sel.innerHTML = '<option value="">— загрузить из файла —</option>' +
      (files || []).map(f => `<option value="${escapeHtml(f.name)}">${escapeHtml(f.name)}</option>`).join('');
    if (currentVal) sel.value = currentVal;
  } catch (_) {}
}

async function loadDshFromFile(filename) {
  if (!filename) return;
  try {
    const r = await fetch(API + '/data/' + encodeURIComponent(filename));
    const data = await r.json();
    const items = (data?.value?.items || data?.items || data?.content || []);
    const flattened = items.map(flattenItem);
    setDashboardContent(flattened);
    renderDashboard();
    const info = el('dsh-source-info');
    if (info) info.textContent = `Файл: ${filename}, записей: ${flattened.length}`;
  } catch (e) {
    console.error('Load dashboard file', e);
    setDashboardContent([]);
    renderDashboard();
  }
}

function renderCompanyFilters() {
  const container = el('company-btns');
  if (!container) return;
  
  if (!window._emplCompanies || window._emplCompanies.length === 0) {
    container.innerHTML = '<button type="button" class="btn btn-company" disabled>Нет данных о компаниях</button>';
    return;
  }
  
  container.innerHTML = window._emplCompanies
    .map(c => `<button type="button" class="btn btn-company" data-company="${escapeHtml(c)}">${escapeHtml(c)}</button>`)
    .join('');
  
  container.querySelectorAll('.btn-company').forEach(btn => {
    btn.addEventListener('click', () => {
      filterCompany = btn.getAttribute('data-company');
      container.querySelectorAll('.btn-company').forEach(b => b.classList.remove('btn-company-active'));
      btn.classList.add('btn-company-active');
      if (viewMode === 'stats') renderTable(null);
    });
  });
}

function renderTable(apiResult) {
  const thead = el('data-thead');
  const tbody = el('data-tbody');
  const empty = el('data-empty');
  const paginationBar = el('pagination-bar');

  if (!thead || !tbody) return;

  if (apiResult != null) {
    let items = [];
    
    if (apiResult.value?.items) {
      items = apiResult.value.items;
    } else if (apiResult.items) {
      items = apiResult.items;
    } else if (Array.isArray(apiResult.content)) {
      items = apiResult.content;
    } else if (Array.isArray(apiResult)) {
      items = apiResult;
    }
    
    window._lastTableData = items.map(flattenItem);
    tableCurrentPage = 1;
  }

  const content = window._lastTableData || [];
  
  if (!content || content.length === 0) {
    thead.innerHTML = '';
    tbody.innerHTML = '<tr><td colspan="20" style="text-align: center; padding: 40px; font-size: 16px; color: #666;">📭 Нет данных для отображения</td></tr>';
    if (empty) empty.style.display = 'block';
    if (paginationBar) paginationBar.style.display = 'flex';
    updatePaginationUI(0, 0);
    return;
  }
  
  if (empty) empty.style.display = 'none';

  if (viewMode === 'stats') {
    let contentByShift = content;
    if (shiftFilter === 'day') {
      contentByShift = content.filter(row => isDayShift(row.operationCompletedAt));
    } else if (shiftFilter === 'night') {
      contentByShift = content.filter(row => isNightShift(row.operationCompletedAt));
    }
    
    let stats = computeStats(contentByShift);
    
    if (filterCompany === '__none__') {
      stats = stats.filter(r => !r.company || r.company === '—');
    } else if (filterCompany !== '__all__') {
      stats = stats.filter(r => r.company === filterCompany);
    }
    
    thead.innerHTML = '<tr><th>Компания</th><th>ФИО</th><th>СЗ</th><th>ХР</th><th>КДК</th><th>ШТ</th><th>Простои >10 мин</th><th>ЕО</th></tr>';
    
    if (stats.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px; font-size: 16px; color: #666;">🔍 Нет данных по выбранным фильтрам</td></tr>';
    } else {
      tbody.innerHTML = stats.map(row => {
        return `<tr>
          <td>${escapeHtml(row.company || '—')}</td>
          <td><strong>${escapeHtml(row.fio)}</strong></td>
          <td class="number-cell">${row.sz || 0}</td>
          <td class="number-cell">${row.hr || 0}</td>
          <td class="number-cell">${row.kdk || 0}</td>
          <td class="number-cell">${row.pieces || 0}</td>
          <td>${escapeHtml(row.idlesStr)}</td>
          <td class="number-cell">${row.eo || 0}</td>
        </tr>`;
      }).join('');
    }
    
    if (paginationBar) paginationBar.style.display = 'none';
    return;
  }

  if (paginationBar) paginationBar.style.display = 'flex';

  const first = content[0];
  const keys = Object.keys(first);
  
  thead.innerHTML = '<tr>' + keys.map(k => 
    `<th>${escapeHtml(TABLE_HEADERS[k] || k)}</th>`
  ).join('') + '</tr>';

  const filterVal = el('filter')?.value || '';
  const lower = filterVal.toLowerCase();
  
  let filtered = content;
  if (filterVal) {
    filtered = content.filter(row => 
      keys.some(k => String(row[k] ?? '').toLowerCase().includes(lower))
    );
  }
  
  const totalRows = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / TABLE_PAGE_SIZE));
  
  if (tableCurrentPage > totalPages) tableCurrentPage = totalPages;
  
  const start = (tableCurrentPage - 1) * TABLE_PAGE_SIZE;
  const pageRows = filtered.slice(start, start + TABLE_PAGE_SIZE);

  tbody.innerHTML = pageRows.map(row => 
    '<tr>' + keys.map(k => 
      `<td>${escapeHtml(formatCell(row[k]))}</td>`
    ).join('') + '</tr>'
  ).join('');

  updatePaginationUI(totalRows, totalPages);
}

function setViewMode(mode) {
  if (viewMode === mode) return;
  viewMode = mode;
  
  const btnStats = el('btn-view-stats');
  const btnFull = el('btn-view-full');
  const filtersEl = el('company-filters');
  
  if (btnStats) btnStats.classList.toggle('btn-tab-active', mode === 'stats');
  if (btnFull) btnFull.classList.toggle('btn-tab-active', mode === 'full');
  if (filtersEl) filtersEl.classList.toggle('company-filters-visible', mode === 'stats');
  
  renderTable(null);
}

function updatePaginationUI(totalRows, totalPages) {
  const info = el('pagination-info');
  const pageNum = el('page-num');
  const firstBtn = el('btn-page-first');
  const prevBtn = el('btn-page-prev');
  const nextBtn = el('btn-page-next');
  const lastBtn = el('btn-page-last');
  
  if (!info || !pageNum) return;
  
  if (totalRows === 0) {
    info.textContent = '—';
    pageNum.textContent = '0';
    if (firstBtn) firstBtn.disabled = true;
    if (prevBtn) prevBtn.disabled = true;
    if (nextBtn) nextBtn.disabled = true;
    if (lastBtn) lastBtn.disabled = true;
    return;
  }
  
  const from = (tableCurrentPage - 1) * TABLE_PAGE_SIZE + 1;
  const to = Math.min(tableCurrentPage * TABLE_PAGE_SIZE, totalRows);
  info.textContent = `Строки ${from}—${to} из ${totalRows}`;
  pageNum.textContent = `${tableCurrentPage} / ${totalPages}`;
  
  if (firstBtn) firstBtn.disabled = tableCurrentPage <= 1;
  if (prevBtn) prevBtn.disabled = tableCurrentPage <= 1;
  if (nextBtn) nextBtn.disabled = tableCurrentPage >= totalPages;
  if (lastBtn) lastBtn.disabled = tableCurrentPage >= totalPages;
}

function goToPage(page) {
  const content = window._lastTableData || [];
  if (content.length === 0) return;
  
  const filterVal = el('filter')?.value || '';
  const keys = content[0] ? Object.keys(content[0]) : [];
  const lower = filterVal.toLowerCase();
  
  const filtered = filterVal
    ? content.filter(row => keys.some(k => String(row[k] ?? '').toLowerCase().includes(lower)))
    : content;
  
  const totalPages = Math.max(1, Math.ceil(filtered.length / TABLE_PAGE_SIZE));
  const next = Math.max(1, Math.min(page, totalPages));
  
  if (next === tableCurrentPage) return;
  
  tableCurrentPage = next;
  renderTable(null);
}

function formatCell(v) {
  if (v == null) return '—';
  if (typeof v === 'object') return JSON.stringify(v);
  if (typeof v === 'boolean') return v ? 'Да' : 'Нет';
  if (typeof v === 'number') return v.toLocaleString('ru-RU');
  return String(v);
}

function escapeHtml(s) {
  if (s === null || s === undefined) return '';
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function addHistoryItem(result) {
  const list = el('history-list');
  if (!list) return;
  
  const li = document.createElement('li');
  const total = result.data?.value?.total ?? result.total ?? result.count ?? 0;
  li.textContent = `${formatDate(new Date())} — на странице: ${result.count ?? 0}, всего: ${total}, ${result.duration ?? 0} мс`;
  list.insertBefore(li, list.firstChild);
  
  while (list.children.length > 30) {
    if (list.lastChild) list.removeChild(list.lastChild);
  }
}

async function startSchedule() {
  const interval = getInterval();
  try {
    await fetch(API + '/config', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ intervalMinutes: interval, token: getToken() || undefined }),
    });
    const r = await fetch(API + '/schedule/start', { method: 'POST' });
    const data = await r.json();
    if (data.ok) {
      setStatus('Статус: Автоопрос включён', 'success', data.message);
    } else {
      setStatus('Ошибка: ' + (data.message || 'не удалось запустить'), 'error', '');
    }
  } catch (e) {
    setStatus('Ошибка', 'error', e.message);
  }
}

async function stopSchedule() {
  try {
    const r = await fetch(API + '/schedule/stop', { method: 'POST' });
    const data = await r.json();
    setStatus('Статус: Остановлено', 'pending', data.message || '');
  } catch (e) {
    setStatus('Ошибка', 'error', e.message);
  }
}

async function saveToken() {
  const token = getToken();
  const cookie = getCookie();
  if (!token && !cookie) return;
  
  try {
    await fetch(API + '/config', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: token || undefined, cookie: cookie || undefined }),
    });
    setStatus('Настройки сохранены', 'success', '');
  } catch (e) {
    setStatus('Ошибка сохранения', 'error', e.message);
  }
}

function setTodayRange() {
  const now = new Date();
  const today = now.toISOString().slice(0, 10);
  const fromEl = el('date-from');
  const toEl = el('date-to');
  
  if (fromEl) fromEl.value = today + 'T00:00:00';
  if (toEl) toEl.value = today + 'T23:59:59';
  
  updateDateDisplays();
}

const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
let datePickerState = { open: false, viewYear: new Date().getFullYear(), viewMonth: new Date().getMonth(), rangeFrom: null, rangeTo: null };

function openDateDropdown() {
  const fromEl = el('date-from');
  const toEl = el('date-to');
  const timeFromEl = el('time-from');
  const timeToEl = el('time-to');
  const dateDropdown = el('date-dropdown');
  
  if (!fromEl || !toEl || !dateDropdown) return;
  
  const from = fromEl.value;
  const to = toEl.value;
  const dFrom = from ? new Date(from) : new Date();
  const dTo = to ? new Date(to) : new Date();
  
  datePickerState.viewYear = dFrom.getFullYear();
  datePickerState.viewMonth = dFrom.getMonth();
  datePickerState.rangeFrom = from ? new Date(dFrom.getFullYear(), dFrom.getMonth(), dFrom.getDate()) : null;
  datePickerState.rangeTo = to ? new Date(dTo.getFullYear(), dTo.getMonth(), dTo.getDate()) : null;
  
  if (datePickerState.rangeFrom) datePickerState.rangeFrom.setHours(0, 0, 0, 0);
  if (datePickerState.rangeTo) datePickerState.rangeTo.setHours(0, 0, 0, 0);
  
  if (timeFromEl) timeFromEl.value = from ? String(from).slice(11, 16) : '00:00';
  if (timeToEl) timeToEl.value = to ? String(to).slice(11, 16) : '23:59';
  
  dateDropdown.classList.add('is-open');
  dateDropdown.setAttribute('aria-hidden', 'false');
  
  fillCalendarSelects();
  renderCalendarDays();
  positionDateDropdown();
}

function positionDateDropdown() {
  const trigger = el('date-range-trigger');
  const dd = el('date-dropdown');
  if (!trigger || !dd) return;
  const rect = trigger.getBoundingClientRect();
  dd.style.position = 'fixed';
  dd.style.left = rect.left + 'px';
  dd.style.top = (rect.bottom + 4) + 'px';
}

function closeDateDropdown() {
  datePickerState.open = false;
  const dd = el('date-dropdown');
  if (dd) {
    dd.classList.remove('is-open');
    dd.setAttribute('aria-hidden', 'true');
  }
}

function fillCalendarSelects() {
  const monthSelect = el('cal-month');
  const yearSelect = el('cal-year');
  if (!monthSelect || !yearSelect) return;
  
  const { viewYear, viewMonth } = datePickerState;
  monthSelect.innerHTML = MONTHS.map((name, i) => `<option value="${i}"${i === viewMonth ? ' selected' : ''}>${name}</option>`).join('');
  
  const yearFrom = viewYear - 5;
  const yearTo = viewYear + 5;
  yearSelect.innerHTML = Array.from({ length: yearTo - yearFrom + 1 }, (_, i) => {
    const y = yearFrom + i;
    return `<option value="${y}"${y === viewYear ? ' selected' : ''}>${y}</option>`;
  }).join('');
}

function renderCalendarDays() {
  const container = el('cal-days');
  if (!container) return;
  
  const { viewYear, viewMonth, rangeFrom, rangeTo } = datePickerState;
  const first = new Date(viewYear, viewMonth, 1);
  const last = new Date(viewYear, viewMonth + 1, 0);
  const startDow = first.getDay() === 0 ? 6 : first.getDay() - 1;
  const daysInMonth = last.getDate();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let html = '';
  const rangeFromTs = rangeFrom ? rangeFrom.getTime() : null;
  const rangeToTs = rangeTo ? rangeTo.getTime() : null;
  
  for (let i = 0; i < startDow; i++) {
    const prevMonth = new Date(viewYear, viewMonth, -startDow + i + 1);
    prevMonth.setHours(0, 0, 0, 0);
    const d = prevMonth.getDate();
    html += `<button type="button" class="calendar-day other-month" data-date="${prevMonth.getFullYear()}-${String(prevMonth.getMonth() + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}">${d}</button>`;
  }
  
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(viewYear, viewMonth, d);
    date.setHours(0, 0, 0, 0);
    const dateStr = `${viewYear}-${String(viewMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dateTs = date.getTime();
    const isToday = dateTs === today.getTime();
    let cls = 'calendar-day';
    if (isToday) cls += ' today';
    if (rangeFromTs !== null && dateTs === rangeFromTs) cls += ' range-from selected';
    else if (rangeToTs !== null && dateTs === rangeToTs) cls += ' range-to selected';
    else if (rangeFromTs !== null && rangeToTs !== null && dateTs > rangeFromTs && dateTs < rangeToTs) cls += ' in-range';
    html += `<button type="button" class="${cls}" data-date="${dateStr}">${d}</button>`;
  }
  
  const totalCells = startDow + daysInMonth;
  const rest = totalCells % 7 === 0 ? 0 : 7 - (totalCells % 7);
  
  for (let i = 0; i < rest; i++) {
    const nextMonth = new Date(viewYear, viewMonth + 1, i + 1);
    nextMonth.setHours(0, 0, 0, 0);
    const d = nextMonth.getDate();
    html += `<button type="button" class="calendar-day other-month" data-date="${nextMonth.getFullYear()}-${String(nextMonth.getMonth() + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}">${d}</button>`;
  }
  
  container.innerHTML = html;
  
  container.querySelectorAll('.calendar-day').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const dateStr = btn.getAttribute('data-date');
      if (!dateStr) return;
      
      const [y, m, d] = dateStr.split('-').map(Number);
      const date = new Date(y, m - 1, d);
      date.setHours(0, 0, 0, 0);
      const hasFullRange = datePickerState.rangeFrom && datePickerState.rangeTo;
      
      if (!datePickerState.rangeFrom || hasFullRange) {
        datePickerState.rangeFrom = new Date(date.getTime());
        datePickerState.rangeTo = null;
      } else {
        if (date.getTime() < datePickerState.rangeFrom.getTime()) {
          datePickerState.rangeTo = new Date(datePickerState.rangeFrom.getTime());
          datePickerState.rangeFrom = new Date(date.getTime());
        } else {
          datePickerState.rangeTo = new Date(date.getTime());
        }
      }
      renderCalendarDays();
      updateDateDisplaysFromPicker();
    });
  });
}

function parseTime(s) {
  if (!s) return null;
  const m = s.trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  if (h < 0 || h > 23 || min < 0 || min > 59) return null;
  return `${String(h).padStart(2, '0')}:${String(min).padStart(2, '0')}`;
}

function getPageSize() {
  const sizeEl = document.querySelector('input[name="page-size"]:checked');
  return sizeEl ? parseInt(sizeEl.value, 10) : 100;
}

function updateProgress(current, total, status) {
  const container = el('progress-container');
  const statusEl = el('progress-status');
  const percentEl = el('progress-percent');
  const fillEl = el('progress-fill');
  const detailEl = el('progress-detail');
  
  if (!container || !statusEl || !percentEl || !fillEl) return;
  
  if (current === 0 && total === 0) {
    container.style.display = 'none';
    return;
  }
  
  container.style.display = 'block';
  
  const percent = total > 0 ? Math.round((current / total) * 100) : 0;
  
  statusEl.textContent = status || 'Загрузка...';
  percentEl.textContent = `${percent}%`;
  fillEl.style.width = `${percent}%`;
  
  if (detailEl) {
    detailEl.textContent = `Загружено: ${current.toLocaleString()} / ${total.toLocaleString()} записей`;
  }
}

function applyDateRange() {
  const { rangeFrom, rangeTo } = datePickerState;
  const timeFrom = parseTime(el('time-from')?.value) || '00:00';
  const timeTo = parseTime(el('time-to')?.value) || '23:59';
  const fromDate = rangeFrom ? new Date(rangeFrom.getTime()) : new Date();
  const toDate = rangeTo ? new Date(rangeTo.getTime()) : (rangeFrom ? new Date(rangeFrom.getTime()) : new Date());
  const pad = n => String(n).padStart(2, '0');
  
  const y = fromDate.getFullYear(), m = fromDate.getMonth(), d = fromDate.getDate();
  const fromValue = `${y}-${pad(m + 1)}-${pad(d)}T${timeFrom}:00`;
  
  const y2 = toDate.getFullYear(), m2 = toDate.getMonth(), d2 = toDate.getDate();
  const toValue = `${y2}-${pad(m2 + 1)}-${pad(d2)}T${timeTo}:00`;
  
  const fromInput = el('date-from');
  const toInput = el('date-to');
  if (fromInput) fromInput.value = fromValue;
  if (toInput) toInput.value = toValue;
  
  const fromDisplay = document.getElementById('date-from-display');
  const toDisplay = document.getElementById('date-to-display');
  if (fromDisplay) fromDisplay.textContent = formatDateTimeForDisplay(fromValue);
  if (toDisplay) toDisplay.textContent = formatDateTimeForDisplay(toValue);
  
  closeDateDropdown();
}

function resetDateRange() {
  setTodayRange();
  closeDateDropdown();
}

function exportCsv(mode) {
  const content = window._lastTableData || [];
  if (!content.length) {
    alert('Нет данных для экспорта');
    return;
  }

  let rows = [];
  let suffix = '';

  if (mode === 'stats') {
    const contentByShift = shiftFilter === 'day'
      ? content.filter(row => isDayShift(row.operationCompletedAt))
      : content.filter(row => isNightShift(row.operationCompletedAt));
    let stats = computeStats(contentByShift);
    
    if (filterCompany === '__none__') {
      stats = stats.filter(r => !r.company || r.company === '—');
    } else if (filterCompany !== '__all__') {
      stats = stats.filter(r => r.company === filterCompany);
    }

    rows.push(['Компания', 'ФИО', 'СЗ', 'ХР', 'КДК', 'ШТ', 'Простои >10 мин', 'ЕО']);
    
    stats.forEach(row => {
      rows.push([
        row.company || '—',
        row.fio || '—',
        row.sz || 0,
        row.hr || 0,
        row.kdk || 0,
        row.pieces || 0,
        row.idlesStr || '—',
        row.eo || 0,
      ]);
    });
    
    suffix = 'stats';
  } else {
    const keys = Object.keys(content[0]);
    rows.push(keys.map(k => TABLE_HEADERS[k] || k));
    
    content.forEach(row => {
      rows.push(keys.map(k => row[k] ?? '—'));
    });
    
    suffix = 'full';
  }

  const csvRows = rows.map(row => 
    row.map(cell => {
      const str = String(cell ?? '');
      if (str.includes(';') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(';')
  );
  
  const csv = csvRows.join('\r\n');
  
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `samokat_export_${suffix}_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ==================== ИНИЦИАЛИЗАЦИЯ ====================

setTimeout(() => {
  console.log('Экстренная инициализация авторизации');
  initAuthHandlers();
}, 500);

document.addEventListener('DOMContentLoaded', () => {
  
  console.log('DOM загружен, инициализация...');
  
  initAuthHandlers();
  loadTokenFromConfig();
  loadConnectionMode();
  
  el('btn-mode-office')?.addEventListener('click', () => setVpnMode(false));
  el('btn-mode-vpn')?.addEventListener('click', () => setVpnMode(true));
  el('btn-fetch')?.addEventListener('click', fetchDataNow);
  el('btn-start')?.addEventListener('click', startSchedule);
  el('btn-stop')?.addEventListener('click', stopSchedule);
  el('btn-save-token')?.addEventListener('click', saveToken);
  el('btn-set-today')?.addEventListener('click', setTodayRange);
  el('btn-mode-server')?.addEventListener('click', () => setConnectionMode('server'));
  el('btn-mode-browser')?.addEventListener('click', () => setConnectionMode('browser'));

  el('date-range-trigger')?.addEventListener('click', (e) => {
    e.stopPropagation();
    if (el('date-dropdown')?.classList.contains('is-open')) closeDateDropdown();
    else openDateDropdown();
  });

  el('date-dropdown')?.addEventListener('click', (e) => e.stopPropagation());

  el('cal-prev')?.addEventListener('click', () => {
    if (datePickerState.viewMonth === 0) {
      datePickerState.viewMonth = 11;
      datePickerState.viewYear--;
    } else {
      datePickerState.viewMonth--;
    }
    fillCalendarSelects();
    renderCalendarDays();
  });

  el('cal-next')?.addEventListener('click', () => {
    if (datePickerState.viewMonth === 11) {
      datePickerState.viewMonth = 0;
      datePickerState.viewYear++;
    } else {
      datePickerState.viewMonth++;
    }
    fillCalendarSelects();
    renderCalendarDays();
  });

  el('cal-month')?.addEventListener('change', () => {
    datePickerState.viewMonth = parseInt(el('cal-month')?.value, 10);
    renderCalendarDays();
  });

  el('cal-year')?.addEventListener('change', () => {
    datePickerState.viewYear = parseInt(el('cal-year')?.value, 10);
    renderCalendarDays();
  });

  el('date-apply')?.addEventListener('click', applyDateRange);
  el('date-reset')?.addEventListener('click', resetDateRange);
  
  el('time-from')?.addEventListener('input', () => { 
    if (el('date-dropdown')?.classList.contains('is-open')) updateDateDisplaysFromPicker(); 
  });
  
  el('time-to')?.addEventListener('input', () => { 
    if (el('date-dropdown')?.classList.contains('is-open')) updateDateDisplaysFromPicker(); 
  });

  document.querySelectorAll('.btn-shift').forEach(btn => {
    btn.addEventListener('click', () => {
      shiftFilter = btn.getAttribute('data-shift');
      document.querySelectorAll('.btn-shift').forEach(b => b.classList.remove('btn-shift-active'));
      btn.classList.add('btn-shift-active');
      if (viewMode === 'stats') renderTable(null);
    });
  });

  document.addEventListener('click', (e) => {
    const dd = el('date-dropdown');
    if (!dd || !dd.classList.contains('is-open')) return;
    const insideDropdown = e.target.closest('#date-dropdown');
    const insideTrigger = e.target.closest('#date-range-trigger');
    if (!insideDropdown && !insideTrigger) closeDateDropdown();
  });

  el('btn-export-csv')?.addEventListener('click', () => {
    const content = window._lastTableData || [];
    if (!content.length) {
      alert('Нет данных для экспорта');
      return;
    }
    
    const inStats = viewMode === 'stats';
    const msg = inStats
      ? 'Что выгружать?\nOK — статистику.\nОтмена — полную таблицу.'
      : 'Что выгружать?\nOK — полную таблицу.\nОтмена — статистику.';
    const ok = window.confirm(msg);
    const mode = inStats ? (ok ? 'stats' : 'full') : (ok ? 'full' : 'stats');

    if (mode === 'stats') {
      setViewMode('stats');
    } else {
      setViewMode('full');
    }

    exportCsv(mode);
  });

  el('btn-view-stats')?.addEventListener('click', () => setViewMode('stats'));
  el('btn-view-full')?.addEventListener('click', () => setViewMode('full'));

  el('filter')?.addEventListener('input', () => {
    if (!window._lastTableData) return;
    tableCurrentPage = 1;
    renderTable(null);
  });

  el('btn-page-first')?.addEventListener('click', () => goToPage(1));
  el('btn-page-prev')?.addEventListener('click', () => goToPage(tableCurrentPage - 1));
  el('btn-page-next')?.addEventListener('click', () => goToPage(tableCurrentPage + 1));
  el('btn-page-last')?.addEventListener('click', () => {
    const content = window._lastTableData || [];
    if (!content.length) return;
    const totalPages = Math.ceil(content.length / TABLE_PAGE_SIZE);
    goToPage(totalPages);
  });

  if ('Notification' in window && Notification.permission === 'default') {
    Notification.requestPermission();
  }

  fetchStatus();
  setInterval(fetchStatus, 10000);
  setTodayRange();
  updateDateDisplays();

  loadEmpl().then(() => {
    renderCompanyFilters();
    const filtersEl = el('company-filters');
    if (filtersEl) filtersEl.classList.toggle('company-filters-visible', viewMode === 'stats');
  });

  // Маршрутизация по hash: #/dsh — дашборд
  window.addEventListener('hashchange', () => showPage(getCurrentPage()));
  showPage(getCurrentPage());

  document.querySelectorAll('.main-nav .nav-link').forEach(link => {
    link.addEventListener('click', (e) => {
      const h = (link.getAttribute('data-hash') ?? link.getAttribute('href') ?? '').trim();
      if (h === '' || h === '#') {
        location.hash = '';
        showPage('main');
        e.preventDefault();
      } else {
        location.hash = h;
        e.preventDefault();
      }
    });
  });

  el('dsh-use-current')?.addEventListener('click', () => {
    _dshContent = null;
    setDashboardContent(window._lastTableData || []);
    const sel = el('dsh-history-file');
    if (sel) sel.value = '';
    renderDashboard();
  });

  el('dsh-history-file')?.addEventListener('change', function () {
    const filename = this.value;
    if (filename) loadDshFromFile(filename);
  });

  el('dsh-btn-stats')?.addEventListener('click', () => {
    dshViewMode = 'stats';
    document.querySelectorAll('.dsh-toggle-btn').forEach(b => b.classList.remove('dsh-toggle-active'));
    el('dsh-btn-stats')?.classList.add('dsh-toggle-active');
    renderDashboard();
  });
  el('dsh-btn-hours')?.addEventListener('click', () => {
    dshViewMode = 'hours';
    document.querySelectorAll('.dsh-toggle-btn').forEach(b => b.classList.remove('dsh-toggle-active'));
    el('dsh-btn-hours')?.classList.add('dsh-toggle-active');
    renderDashboard();
  });

  el('dsh-btn-export-png')?.addEventListener('click', exportTableToPng);
  el('dsh-btn-export-csv')?.addEventListener('click', exportDashboardCsv);

});

/** Сохранение таблицы дашборда в PNG в высоком качестве (html2canvas). */
async function exportTableToPng() {
  const table = document.getElementById('dsh-unified-table');
  const emptyEl = el('dsh-empty');
  if (!table || !table.rows.length) {
    alert('Нет данных для экспорта.');
    return;
  }
  if (emptyEl && emptyEl.style.display !== 'none') {
    alert('Нет данных для экспорта.');
    return;
  }
  if (typeof html2canvas === 'undefined') {
    alert('Библиотека html2canvas не загружена. Проверьте подключение к интернету.');
    return;
  }
  const btn = el('dsh-btn-export-png');
  if (btn) btn.disabled = true;
  try {
    const canvas = await html2canvas(table, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#0f1419',
      logging: false,
    });
    const dataUrl = canvas.toDataURL('image/png');
    const name = `owms-table-${new Date().toISOString().slice(0, 10)}-${Date.now()}.png`;
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = name;
    a.click();
  } catch (e) {
    console.error('exportTableToPng', e);
    alert('Ошибка при создании PNG: ' + (e.message || e));
  } finally {
    if (btn) btn.disabled = false;
  }
}

/** Выгрузка текущей таблицы дашборда в CSV (статистика или СЗ по часам). */
function exportDashboardCsv() {
  const content = getDashboardContent();
  if (!content || !content.length) {
    alert('Нет данных для выгрузки. Загрузите данные на главной или выберите файл.');
    return;
  }
  let rows = [];
  let suffix = '';
  if (dshViewMode === 'stats') {
    const stats = computeStats(content);
    stats.sort((a, b) => {
      const companyA = (a.company || '—').localeCompare(b.company || '—');
      if (companyA !== 0) return companyA;
      return (b.sz || 0) - (a.sz || 0);
    });
    rows.push(['Компания', 'ФИО', 'СЗ', 'ХР', 'КДК', 'ШТ', 'Простои >10 мин', 'ЕО']);
    stats.forEach(row => {
      rows.push([
        row.company || '—',
        row.fio || '—',
        row.sz || 0,
        row.hr || 0,
        row.kdk || 0,
        row.pieces || 0,
        row.idlesStr || '—',
        row.eo || 0,
      ]);
    });
    suffix = 'stats';
  } else {
    const { fios, hours, getSz } = computeSzByHour(content);
    rows.push(['Компания', 'ФИО', ...hours.map(String), 'Σ']);
    fios.forEach(fio => {
      const company = getCompanyForFio(fio) || '—';
      const rowVals = hours.map(h => getSz(fio, h));
      const total = rowVals.reduce((a, b) => a + b, 0);
      rows.push([company, fio, ...rowVals, total]);
    });
    suffix = 'hours';
  }
  const csvRows = rows.map(row =>
    row.map(cell => {
      const str = String(cell ?? '');
      if (str.includes(';') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(';')
  );
  const csv = csvRows.join('\r\n');
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `owms_dashboard_${suffix}_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}