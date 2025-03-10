#!/bin/bash

# Activar entorno conda
source ~/anaconda3/etc/profile.d/conda.sh
conda activate sat-laboral

# Ir al directorio del script (Cambiar la ruta)
cd "$(dirname "$0")"

# Ejecutar el archivo Python
python Generar_informe.py
