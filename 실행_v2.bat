@echo off
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
set "DART_UI_PY=python"
powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 8772 -State Listen -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"
if not errorlevel 1 (
  start "" "http://127.0.0.1:8772/"
  exit /b 0
)
"%DART_UI_PY%" server_v2.py --port 8772
pause
