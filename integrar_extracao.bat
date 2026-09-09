@echo off
setlocal
cd /d "%~dp0_scripts" || exit /b 1
where py >nul 2>&1
if %errorlevel%==0 (
  py -3 integrar_extracao_planilha.py
) else (
  python integrar_extracao_planilha.py
)
pause
