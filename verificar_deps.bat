@echo off
setlocal
cd /d "%~dp0_scripts" || (
  echo Nao foi possivel acessar a pasta _scripts.
  pause
  exit /b 1
)

where py >nul 2>&1
if %errorlevel%==0 (
  py -3 verificar_deps.py
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo Python nao foi encontrado. Instale Python 3 e marque "Add Python to PATH".
    pause
    exit /b 1
  )
  python verificar_deps.py
)

echo.
set /p resposta=Instalar as bibliotecas Python que faltam? (S/N):
if /I "%resposta%"=="S" (
  where py >nul 2>&1
  if %errorlevel%==0 (
    py -3 verificar_deps.py --instalar
  ) else (
    python verificar_deps.py --instalar
  )
)
echo.
pause
