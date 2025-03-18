@echo off
REM Obtener el nombre del entorno desde environment.yml (En Windows, se asume el nombre)
set ENV_NAME=sat-laboral

:: Verificar si Conda está instalado
where conda >nul 2>nul
if %errorlevel% neq 0 (
    echo Conda no está instalado. Instálalo y vuelve a intentarlo.
    exit /b 1
)

:: Verificar si el entorno ya existe
conda env list | findstr /C:"%ENV_NAME%" >nul
if %errorlevel% equ 0 (
    echo El entorno '%ENV_NAME%' ya existe. Actualizando paquetes...
    conda env update --name "%ENV_NAME%" --file environment.yml --prune
    echo Entorno '%ENV_NAME%' actualizado correctamente.
) else (
    echo Creando el entorno '%ENV_NAME%' desde environment.yml...
    conda env create -f environment.yml
    if %errorlevel% equ 0 (
        echo Entorno '%ENV_NAME%' creado con éxito.
    ) else (
        echo Hubo un error al crear el entorno.
        exit /b 1
    )
)
