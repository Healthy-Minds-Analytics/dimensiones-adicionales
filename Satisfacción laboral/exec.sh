#!/usr/bin/env bash

echo "Detectando rutas..."

# 1) DETECTAR EL DIRECTORIO DONDE ESTÁ ESTE .SH
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 2) DEFINIR NOMBRE DEL ENTORNO, RUTAS Y ARCHIVOS
ENV_NAME="sat-laboral"
CONDA_PATH="$HOME/anaconda3"
SCRIPT_FILE="$SCRIPT_DIR/Generar_informe.py"
ENV_FILE="$SCRIPT_DIR/environment.yml"

# 3) VERIFICAR SI CONDA ESTÁ INSTALADO (SIN DEPENDER DEL PATH)
if [ ! -f "$CONDA_PATH/bin/conda" ]; then
    echo
    echo "[ERROR] No se ha encontrado conda en: $CONDA_PATH/bin/conda"
    echo "Debes asegurarte de haber instalado Anaconda o Miniconda en la carpeta de tu usuario."
    exit 1
fi

echo
echo "Anaconda detectado correctamente en: $CONDA_PATH/bin/conda"

# 4) INICIALIZAR CONDA en esta shell
if [ -f "$CONDA_PATH/etc/profile.d/conda.sh" ]; then
    # método recomendado
    source "$CONDA_PATH/etc/profile.d/conda.sh"
else
    # fallback a script legacy (raro en Mac pero por si acaso)
    source "$CONDA_PATH/bin/activate"
fi

# 5) COMPROBAR SI EXISTE EL ENTORNO sat-laboral
echo
if conda env list | grep -qE "^${ENV_NAME}[[:space:]]"; then
    echo "El entorno \"$ENV_NAME\" ya existe. Se procede a actualizarlo..."
    conda env update --name "$ENV_NAME" --file "$ENV_FILE" --prune
    if [ $? -ne 0 ]; then
        echo
        echo "[ERROR] Hubo un problema al actualizar el entorno \"$ENV_NAME\"."
        exit 1
    fi
    echo
    echo "Entorno \"$ENV_NAME\" actualizado correctamente."
else
    echo "El entorno \"$ENV_NAME\" no existe. Se procedera a crearlo..."
    conda env create -y -f "$ENV_FILE"
    if [ $? -ne 0 ]; then
        echo
        echo "[ERROR] Hubo un problema al crear el entorno \"$ENV_NAME\"."
        exit 1
    fi
    echo
    echo "Entorno \"$ENV_NAME\" creado correctamente."
fi

# 6) ACTIVAR EL ENTORNO sat-laboral
source "$CONDA_PATH/bin/activate" "$ENV_NAME"
if [ $? -ne 0 ]; then
    echo
    echo "[ERROR] No se pudo activar el entorno \"$ENV_NAME\"."
    exit 1
fi

# 7) COMPROBAR QUE Generar_informe.py EXISTA
if [ ! -f "$SCRIPT_FILE" ]; then
    echo
    echo "[ERROR] No se ha encontrado el archivo Generar_informe.py en la ruta:"
    echo "$SCRIPT_FILE"
    exit 1
fi

# 8) EJECUTAR EL SCRIPT Generar_informe.py
echo
echo "Ejecutando \"$SCRIPT_FILE\" con Python del entorno \"$ENV_NAME\"..."
python "$SCRIPT_FILE"
if [ $? -ne 0 ]; then
    echo
    echo "[ERROR] Hubo un problema al ejecutar Generar_informe.py"
    exit 1
fi

echo
echo "El script Generar_informe.py se ha ejecutado correctamente."

# 9) MENSAJE FINAL: NO SE CIERRA LA VENTANA TRAS PULSAR TECLA
echo
read -n 1 -s -r -p "Presione cualquier tecla para salir..."
echo
