@echo off
chcp 65001 >nul
echo.
echo ================================
echo   WAPSA - Actualizar Dashboard
echo ================================
echo.

:: Verificar que existe algun archivo fuente en input/
dir /b "input\WAPSA Team - Epics*.xlsx" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontro ningun archivo fuente en la carpeta input/
    echo         Copia el xlsx exportado de Azure DevOps ahi primero.
    echo.
    pause
    exit /b 1
)

:: Ejecutar script de conversion
echo [1/3] Convirtiendo datos...
python convert.py
if errorlevel 1 (
    echo [ERROR] Fallo la conversion. Revisa el archivo fuente.
    echo.
    pause
    exit /b 1
)

:: Git add + commit con fecha de hoy
echo.
echo [2/3] Preparando commit...
git add data.xlsx

:: Obtener fecha actual
for /f "tokens=1-3 delims=/" %%a in ("%date%") do (
    set DIA=%%a
    set MES=%%b
    set ANIO=%%c
)
git commit -m "Actualizar datos - %DIA%-%MES%-%ANIO%"

:: Push
echo.
echo [3/3] Subiendo al servidor...
git push
if errorlevel 1 (
    echo [ERROR] Fallo el push. Verifica tu conexion o credenciales de Git.
    echo.
    pause
    exit /b 1
)

echo.
echo ================================
echo   Listo! Dashboard actualizado.
echo   Espera ~1 minuto en GitHub Pages.
echo ================================
echo.
pause
