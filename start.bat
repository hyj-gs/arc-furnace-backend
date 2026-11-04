@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Arc Furnace Recorder

echo.
echo ========================================
echo   Arc Furnace Recorder - Multi-User
echo ========================================
echo.

where node >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js not found!
    echo.
    echo Please install Node.js from https://nodejs.org/
    echo Download LTS version, install it, then reopen this window.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('node -v') do echo [OK] Node.js: %%i

if not exist "node_modules" (
    echo.
    echo [First run] Installing dependencies...
    call npm install --registry=https://registry.npmmirror.com
    if errorlevel 1 (
        echo [ERROR] npm install failed
        pause
        exit /b 1
    )
    echo [OK] Dependencies ready
)

echo.
echo ========================================
echo   Starting...
echo   Open browser: http://localhost:3000
echo ========================================
echo.

node server.js
pause
