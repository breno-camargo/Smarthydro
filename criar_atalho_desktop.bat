@echo off
chcp 65001 > nul
title Criar Atalho na Area de Trabalho — CompaSSS SmartHydro
cls
echo =====================================================================
echo  Criador de Atalho do SmartHydro na Area de Trabalho
echo =====================================================================
echo.

set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if not exist "%PYTHON_EXE%" (
    set PYTHON_EXE=python
)

"%PYTHON_EXE%" -c "import os, win32com.client; shell = win32com.client.Dispatch('WScript.Shell'); desktop = shell.SpecialFolders('Desktop'); lnk = shell.CreateShortCut(os.path.join(desktop, 'SmartHydro - CompaSSS.lnk')); lnk.TargetPath = os.path.abspath(r'%~dp0dist\RelatorioHidrometros.exe'); lnk.WorkingDirectory = os.path.abspath(r'%~dp0dist'); lnk.IconLocation = os.path.abspath(r'%~dp0app_icon.ico') + ',0'; lnk.Description = 'SmartHydro - Medição de Água Praça Pamplona (CompaSSS)'; lnk.save()"

if %ERRORLEVEL% EQU 0 (
    echo =====================================================================
    echo  SUCESSO!
    echo  Atalho 'SmartHydro - CompaSSS' criado com sucesso na sua Area de Trabalho!
    echo =====================================================================
) else (
    echo [ERRO] Falha ao criar o atalho.
)
echo.
pause
