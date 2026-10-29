@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
set "UNIFIER_PYTHON=%~dp0.venv\Scripts\python.exe"
if exist "%UNIFIER_PYTHON%" goto run_venv
python "%~dp0excel_unifier.py" %*
exit /b %errorlevel%
:run_venv
"%UNIFIER_PYTHON%" "%~dp0excel_unifier.py" %*
exit /b %errorlevel%
