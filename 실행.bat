@echo off
cd /d "%~dp0"
set PY=python
where %PY% >nul 2>&1
if errorlevel 1 (
  echo [dart_fs] Python not found: %PY%
  echo [dart_fs] 파이썬을 찾지 못했습니다. 실행.bat 의 PY 경로를 고치세요.
  pause
  exit /b 1
)
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 8771 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { $p = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue; if ($p -and $p.ProcessName -like 'python*') { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } }" >nul 2>&1
timeout /t 1 /nobreak >nul
"%PY%" server.py
pause
