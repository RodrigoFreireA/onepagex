@echo off
title OnePageReport
cd /d "%~dp0"

echo.
echo  Iniciando OnePageReport...
echo.

python server.py

REM Se chegou aqui, o servidor encerrou normalmente (Ctrl+C ou /shutdown)
REM A janela fecha automaticamente apos 2 segundos
if %errorlevel% equ 0 (
    echo.
    echo  [OK] Servidor encerrado.
    timeout /t 2 /nobreak >nul
    exit
)

REM Erro ao iniciar
echo.
echo  [ERRO] Falha ao iniciar. Verifique se o Python esta instalado.
echo  Instale as dependencias com:
echo      pip install pandas openpyxl
echo.
pause
