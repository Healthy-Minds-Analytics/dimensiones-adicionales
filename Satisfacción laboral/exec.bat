@echo off

echo Detectando rutas...

REM 1) DETECTAR EL DIRECTORIO DONDE ESTÁ ESTE .BAT
set SCRIPT_DIR=%~dp0

REM 2) DEFINIR NOMBRE DEL ENTORNO, RUTAS Y ARCHIVOS
set ENV_NAME=sat-laboral
set CONDA_PATH=%USERPROFILE%\anaconda3
set SCRIPT_FILE=%SCRIPT_DIR%Generar_informe.py
set ENV_FILE=%SCRIPT_DIR%environment.yml

REM 3) VERIFICAR SI CONDA ESTÁ INSTALADO (SIN DEPENDER DEL PATH)
if not exist "%CONDA_PATH%\Scripts\conda.exe" (
    echo.
    echo [ERROR] No se ha encontrado conda en: %CONDA_PATH%\Scripts\conda.exe
    echo Debes asegurarte de haber instalado Anaconda o Miniconda en la carpeta de tu usuario.
    goto end
)

echo.
echo Anaconda detectado correctamente en: %CONDA_PATH%\Scripts\conda.exe

REM 4) ACTIVAR BASE E INICIALIZAR CONDA
call "%CONDA_PATH%\Scripts\activate.bat" base

REM 5) COMPROBAR SI EXISTE EL ENTORNO sat-laboral
echo.
call conda env list | findstr /C:"%ENV_NAME%" >nul
if %errorlevel% equ 0 (
    echo El entorno "%ENV_NAME%" ya existe. Se procede a actualizarlo...
    call conda env update --name "%ENV_NAME%" --file "%ENV_FILE%" --prune
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] Hubo un problema al actualizar el entorno "%ENV_NAME%".
        goto end
    )
    echo.
    echo Entorno "%ENV_NAME%" actualizado correctamente.
) else (
    echo El entorno "%ENV_NAME%" no existe. Se procedera a crearlo...
    call conda env create -y -f "%ENV_FILE%"
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] Hubo un problema al crear el entorno "%ENV_NAME%".
        goto end
    )
    echo.
    echo Entorno "%ENV_NAME%" creado correctamente.
)


REM 6) ACTIVAR EL ENTORNO sat-laboral
call conda activate %ENV_NAME%
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] No se pudo activar el entorno "%ENV_NAME%".
    goto end
)

REM 7) COMPROBAR QUE Generar_informe.py EXISTA
if not exist "%SCRIPT_FILE%" (
    echo.
    echo [ERROR] No se ha encontrado el archivo Generar_informe.py en la ruta:
    echo "%SCRIPT_FILE%"
    goto end
)

REM 8) EJECUTAR EL SCRIPT Generar_informe.py
echo.
echo Ejecutando "%SCRIPT_FILE%" con Python del entorno "%ENV_NAME%"...
python "%SCRIPT_FILE%"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Hubo un problema al ejecutar Generar_informe.py
    goto end
)

echo.
echo El script Generar_informe.py se ha ejecutado correctamente.

:end
REM 9) MENSAJE FINAL: NO SE CIERRA LA VENTANA TRAS PULSAR TECLA, SE QUEDA EN EL CMD
echo.
echo Presione cualquier tecla para salir...
pause
exit