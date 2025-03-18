@echo off
REM Forzar la activación de Conda en Windows y evitar que se cierre la terminal

:: Configurar Conda en la sesión actual (si no está en PATH)
CALL "%USERPROFILE%\anaconda3\Scripts\activate.bat" base
CALL conda init

:: Definir el nombre del entorno
set ENV_NAME=sat-laboral

:: Verificar si Conda está instalado correctamente
conda --version >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Conda no está disponible en el PATH del sistema.
    echo Asegúrate de haber instalado Anaconda o Miniconda correctamente.
    echo Si usas Anaconda, abre 'Anaconda Prompt' y ejecuta este script desde allí.
    echo.
    echo Presiona cualquier tecla para salir...
    pause >nul
    exit /b 1
)

:: Verificar si el entorno ya existe
conda env list | findstr /C:"%ENV_NAME%" >nul
if %errorlevel% equ 0 (
    echo.
    echo El entorno '%ENV_NAME%' ya existe. Actualizando paquetes...
    conda env update --name "%ENV_NAME%" --file environment.yml --prune
    if %errorlevel% neq 0 (
        echo ERROR: Hubo un problema al actualizar el entorno.
    ) else (
        echo Entorno '%ENV_NAME%' actualizado correctamente.
    )
) else (
    echo.
    echo Creando el entorno '%ENV_NAME%' desde environment.yml...
    conda env create -f environment.yml
    if %errorlevel% neq 0 (
        echo ERROR: Hubo un problema al crear el entorno.
    ) else (
        echo Entorno '%ENV_NAME%' creado con éxito.
    )
)

echo.
echo Para activarlo, usa: conda activate %ENV_NAME%
echo.

:: Mantener la terminal abierta en TODOS LOS CASOS
echo Presiona cualquier tecla para salir...
pause >nul
cmd /k
