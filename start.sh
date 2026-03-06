#!/bin/bash
# OWMS — запуск на отдельном порту (Ubuntu 24.04 LTS)
# Проверяет Node.js и зависимости, при необходимости устанавливает, запускает сервер.

set -e
cd "$(dirname "$0")"
SCRIPT_DIR="$(pwd)"
PROJECT_NAME="OWMS Samokat Collector"

# Порт по умолчанию (чтобы не конфликтовать с другим проектом на 3000)
DEFAULT_PORT=3001
if [ -f .env ]; then
  source .env 2>/dev/null || true
fi
export PORT="${PORT:-$DEFAULT_PORT}"

echo "=============================================="
echo "  $PROJECT_NAME"
echo "  Порт: $PORT"
echo "=============================================="

# --- 1. Проверка / установка Node.js (>= 18) ---
check_node() {
  if command -v node >/dev/null 2>&1; then
    local ver
    ver=$(node -v 2>/dev/null | sed 's/^v//' | cut -d. -f1)
    if [ -n "$ver" ] && [ "$ver" -ge 18 ] 2>/dev/null; then
      echo "[OK] Node.js $(node -v)"
      return 0
    fi
  fi
  return 1
}

if ! check_node; then
  echo "[!] Node.js 18+ не найден. Установка..."
  if command -v curl >/dev/null 2>&1; then
    curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
    sudo apt-get install -y nodejs
  else
    sudo apt-get update
    sudo apt-get install -y curl
    curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
    sudo apt-get install -y nodejs
  fi
  if ! check_node; then
    echo "[ОШИБКА] Не удалось установить Node.js 18+."
    exit 1
  fi
fi

# --- 2. npm зависимости ---
if [ ! -d node_modules ] || [ ! -f node_modules/.package-lock.json ]; then
  echo "[*] Установка зависимостей (npm install)..."
  npm install
else
  echo "[*] Проверка зависимостей..."
  npm install --no-audit --no-fund 2>/dev/null || npm install
fi
echo "[OK] Зависимости готовы."

# --- 3. Запуск сервера ---
echo ""
echo "Запуск сервера: http://0.0.0.0:$PORT"
echo "Остановка: Ctrl+C"
echo "=============================================="
exec node backend/server.js
