#!/bin/bash

# Obtener el nombre del entorno desde environment.yml
ENV_NAME=$(grep -m 1 '^name:' environment.yml | awk '{print $2}')

# Verificar si Conda está instalado
if ! command -v conda &> /dev/null
then
    echo "Conda no está instalado. Instálalo y vuelve a intentarlo."
    echo "Presiona ENTER para salir..."
    read -r
    exit 1
fi

# Verificar si el entorno ya existe
if conda env list | grep -q "$ENV_NAME"; then
    echo "El entorno '$ENV_NAME' ya existe. Actualizando paquetes..."
    conda env update --name "$ENV_NAME" --file environment.yml --prune
    if [ $? -eq 0 ]; then
        echo "Entorno '$ENV_NAME' actualizado correctamente."
    else
        echo "Hubo un error al actualizar el entorno."
    fi
else
    echo "Creando el entorno '$ENV_NAME' desde environment.yml..."
    conda env create -f environment.yml
    if [ $? -eq 0 ]; then
        echo "Entorno '$ENV_NAME' creado con éxito."
    else
        echo "Hubo un error al crear el entorno."
    fi
fi

echo "Para activarlo, usa: conda activate $ENV_NAME"

# No cerrar la terminal, esperar input del usuario
echo "Presiona ENTER para salir..."
read -r
