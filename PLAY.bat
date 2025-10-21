@echo off
echo Iniciando consolidacao das planilhas...
echo.
cd /d "%~dp0scripts"
.\.venv\Scripts\python.exe atualizar_planilhas.py
echo.
echo Processo concluido!
pause