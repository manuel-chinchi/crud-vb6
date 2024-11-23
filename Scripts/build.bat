@echo off

REM Ruta del compilador VB6.EXE del sistema
set VB6_COMPILER="C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"

REM Path al proyecto y archivos de recursos
set PROJECT_PATH=%~dp0..\crud_vb6.vbp
set RC_FILE="%~dp0..\Resources\app.rc"
set RES_FILE="%~dp0..\Resources\app.res"

REM Paso 1: Compilo el archivo .rc como un archivo de recurso .res

echo Compilando archivo .rc en .res...
rc.exe %RC_FILE%
if %ERRORLEVEL% neq 0 (
    echo Error al compilar %RC_FILE%.
    pause
    exit /b %ERRORLEVEL%
)

REM Paso 2: Compilo el pryecto .vbp

REM El archivo de proyecto deber apuntar al archivo .res generado automaticamen
REM te en el paso anterior

echo Compilando el proyecto VB6...
%VB6_COMPILER% /make %PROJECT_PATH%
if %ERRORLEVEL% neq 0 (
    echo Error al compilar el proyecto VB6.
    pause
    exit /b %ERRORLEVEL%
)

echo Compilación completada con éxito.
pause
