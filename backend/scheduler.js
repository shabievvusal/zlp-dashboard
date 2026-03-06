const path = require('path');
const fs = require('fs');

const CONFIG_PATH = path.join(__dirname, 'config.json');
const DATA_DIR = path.join(__dirname, 'data');

let intervalId = null;
let lastRun = null;
let fetchHandler = null;

function loadConfig() {
  try {
    const raw = fs.readFileSync(CONFIG_PATH, 'utf8');
    return JSON.parse(raw);
  } catch (e) {
    return { token: '', intervalMinutes: 60, apiUrl: 'https://api.samokat.ru/wmsops-wwh/stocks/changes/search', useVpn: true };
  }
}

function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
}

function setFetchHandler(handler) {
  fetchHandler = handler;
}

function start() {
  if (intervalId) return { ok: true, message: 'Уже запущен' };
  const config = loadConfig();
  if (!config.token || config.token.trim() === '') {
    return { ok: false, message: 'Токен не задан. Укажите Bearer токен в настройках.' };
  }
  const minutes = Math.max(1, parseInt(config.intervalMinutes, 10) || 60);
  ensureDataDir();
  intervalId = setInterval(async () => {
    lastRun = new Date();
    if (fetchHandler) await fetchHandler();
  }, minutes * 60 * 1000);
  return { ok: true, message: `Автоопрос каждые ${minutes} мин` };
}

function stop() {
  if (intervalId) {
    clearInterval(intervalId);
    intervalId = null;
  }
  return { ok: true, message: 'Остановлено' };
}

function isRunning() {
  return !!intervalId;
}

function getLastRun() {
  return lastRun;
}

module.exports = {
  loadConfig,
  ensureDataDir,
  setFetchHandler,
  start,
  stop,
  isRunning,
  getLastRun,
  CONFIG_PATH,
  DATA_DIR,
};
