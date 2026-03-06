@echo off
chcp 65001 >nul 2>&1
title OWMS Samokat Collector

cd /d "%~dp0"

set "PROJECT_NAME=OWMS Samokat Collector"
set "DEFAULT_PORT=3001"

:: Читаем порт из .env если файл существует
if exist .env (
    for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
        set "%%A=%%B"
    )
)
if not defined PORT set "PORT=%DEFAULT_PORT%"

echo ==============================================
echo   %PROJECT_NAME%
echo   Port: %PORT%
echo ==============================================

:: --- 1. Проверка Node.js ---
where node >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js not found.
    echo Install Node.js 18+ from https://nodejs.org/
    pause
    exit /b 1
)

for /f "tokens=1 delims=." %%V in ('node -v 2^>nul') do set "NODE_VER=%%V"
set "NODE_VER=%NODE_VER:v=%"
if %NODE_VER% LSS 18 (
    echo [ERROR] Node.js 18+ required. Current: v%NODE_VER%
    echo Update from https://nodejs.org/
    pause
    exit /b 1
)

for /f "delims=" %%V in ('node -v') do echo [OK] Node.js %%V

:: --- 2. npm dependencies ---
if not exist node_modules (
    echo [*] Installing dependencies...
    call npm install
) else (
    echo [*] Checking dependencies...
    call npm install --no-audit --no-fund 2>nul || call npm install
)
echo [OK] Dependencies ready.

:: --- 3. Start server ---
echo.
echo Server: http://localhost:%PORT%
echo Stop: Ctrl+C
echo ==============================================
node backend/server.js
if errorlevel 1 (
    echo.
    echo [ERROR] Server stopped with error.
    pause
)
