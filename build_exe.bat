@echo off
chcp 65001 > nul
echo =====================================================================
echo  Compilador do Software de Relatorios de Hidrometros (Praca Pamplona)
echo =====================================================================
echo.

set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if not exist "%PYTHON_EXE%" (
    set PYTHON_EXE=python
)

echo [1/3] Verificando instalacao do PyInstaller...
"%PYTHON_EXE%" -m pip install --quiet pyinstaller pyodbc pandas openpyxl python-dateutil pillow tkcalendar babel

echo [2/3] Compilando executavel unico (.exe)...
"%PYTHON_EXE%" -m PyInstaller --noconsole --onefile ^
    --name "RelatorioHidrometros" ^
    --add-data "config.json;." ^
    --add-data "logo_final.png;." ^
    --add-data "gui_logo.png;." ^
    --add-data "modelo_relatorio.xlsx;." ^
    --hidden-import "babel.numbers" ^
    --clean ^
    main.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] Falha durante a compilacao com o PyInstaller.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/3] Copiando arquivos para a pasta dist...
copy config.json dist\config.json > nul
copy logo_final.png dist\logo_final.png > nul
copy gui_logo.png dist\gui_logo.png > nul
copy modelo_relatorio.xlsx dist\modelo_relatorio.xlsx > nul

echo.
echo =====================================================================
echo  SUCESSO! O executavel foi gerado em:
echo  %~dp0dist\RelatorioHidrometros.exe
echo =====================================================================
echo.
pause
