@echo off
echo.
echo   XylemView Pro - Build Script
echo   =============================
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo   [ERROR] Node.js not found.
    echo   Install from: https://nodejs.org/
    pause
    exit /b 1
)
echo   Found Node.js:
node --version
echo.

echo   [1/2] Installing dependencies...
call npm install
if errorlevel 1 (
    echo   [ERROR] npm install failed
    pause
    exit /b 1
)
echo.

echo   [2/2] Building for Windows...
call npx electron-builder --win
echo.

echo   Done! Check the dist\ folder for:
echo     XylemView Pro Setup.exe
echo.
pause
