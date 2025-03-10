@echo off
:: Activar Anaconda y entorno
CALL "C:\Users\rober\anaconda3\Scripts\activate.bat" sat-laboral

:: Ir a la carpeta donde está el script
cd /d "C:\Users\rober\Documents\Dimensiones adicionales\dimensiones-adicionales\Satisfacción laboral"

:: Ejecutar el script Python
python Generar_informe.py

:: Mantener la ventana abierta al finalizar
pause
