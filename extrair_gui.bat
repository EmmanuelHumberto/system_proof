@echo off
setlocal
cd /d "%~dp0_scripts" || (
  echo Nao foi possivel acessar a pasta _scripts.
  pause
  exit /b 1
)

where py >nul 2>&1
if %errorlevel%==0 (
  py -3 extrair_gui.py
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo Python nao foi encontrado. Instale Python 3 e marque "Add Python to PATH".
    pause
    exit /b 1
  )
  python extrair_gui.py
)
if errorlevel 1 pause
