@echo off
chcp 65001 > nul
title Compilador CompaSSS — SmartHydro
cls
echo =====================================================================
echo  CompaSSS — Compilador do Executavel (.exe) SmartHydro
echo =====================================================================
echo.

set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if not exist "%PYTHON_EXE%" (
    set PYTHON_EXE=python
)

echo [1/3] Finalizando processos anteriores do software se estiverem abertos...
taskkill /F /IM RelatorioHidrometros.exe >nul 2>&1

echo [2/3] Compilando executavel unico (.exe) com PyInstaller...
"%PYTHON_EXE%" -m PyInstaller --noconfirm RelatorioHidrometros.spec

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo =====================================================================
    echo [ERRO] Ocorreu uma falha durante a compilacao com o PyInstaller.
    echo =====================================================================
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/3] Sincronizando arquivos complementares na pasta dist...
if exist "config.json" copy /Y config.json dist\config.json > nul
if exist "modelo_email.html" copy /Y modelo_email.html dist\modelo_email.html > nul
if exist "modelo_relatorio.xlsx" copy /Y modelo_relatorio.xlsx dist\modelo_relatorio.xlsx > nul
if exist "app_icon.ico" copy /Y app_icon.ico dist\app_icon.ico > nul
if exist "logo_final.png" copy /Y logo_final.png dist\logo_final.png > nul

echo.
echo =====================================================================
echo  SUCESSO ABSOLUTO!
echo  O executavel oficial da CompaSSS foi gerado em:
echo  %~dp0dist\RelatorioHidrometros.exe
echo =====================================================================
echo.
pause
