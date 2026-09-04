@echo off
rem Gelius Synergy: Excel table -> page -> GitHub Pages.
rem Everything here is latin on purpose: cmd.exe reads .cmd in the OEM codepage
rem and cyrillic literals break parsing before the file even runs.
rem All output and the commit message live in sync.py, which is UTF-8.
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set PY=C:\priemka\.venv\Scripts\python.exe
if not exist "%PY%" set PY=python

"%PY%" sync.py --deploy
if errorlevel 1 (
  echo.
  echo Ostanovleno: stranica i sayt ne izmeneny.
  endlocal & exit /b 1
)
endlocal & exit /b 0
