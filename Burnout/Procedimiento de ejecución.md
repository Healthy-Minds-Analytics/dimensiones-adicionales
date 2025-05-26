# Generación de Informes de CBB

## Descripción

Este proyecto genera automáticamente un informe del CBB en formato Word a partir de las respuestas contenidas en un CSV. El flujo de trabajo es:

1. Detectar el último archivo CSV modificado en la carpeta `Respuestas`.
2. Calcular los índices de cada una de las dimensiones del CBB.
3. Rellenar los marcadores de una plantilla de Word (`Plantillas/plantilla_burnout.docx`).
4. Guardar el informe resultante en `Informes generados`.

Solo es necesario ejecutar el script correspondiente y responder a las preguntas que aparezcan.

## Estructura de carpetas

```
./
├── environment.yml
├── Generar_informe.py
├── exec.bat
├── exec.sh
├── Plantillas/
│   └── plantilla_burnout.docx
├── Respuestas/
│   └── respuestas_YYYYMMDD_HHMM.csv
└── Informes generados/
```

## Prerrequisitos

* **Anaconda** o **Miniconda** instalado en la carpeta de usuario:

  * Windows: `%USERPROFILE%\anaconda3`
  * Mac/Linux: `$HOME/anaconda3`
* El archivo `environment.yml` debe encontrarse en la raíz del proyecto y listar todas las dependencias necesarias (pandas, python-docx, etc.).
* No es necesario crear o actualizar el entorno manualmente: los scripts `exec.bat` y `exec.sh` lo gestionan automáticamente.

## Uso

### Windows

1. Copia el archivo CSV con las respuestas en la carpeta `Respuestas`.
2. En el Explorador de archivos, navega a la raíz del proyecto y haz doble clic en `exec.bat`.
3. Se abrirá automáticamente una ventana de consola; sigue las indicaciones introduciendo:

   * Nombre de la empresa.
   * Número de invitados.
4. Al finalizar, la ventana mostrará un mensaje y permanecerá abierta hasta que pulses una tecla.
5. El informe generado se guardará en `Informes generados`.

### Mac/Linux

1. Copia el archivo CSV con las respuestas en la carpeta `Respuestas`.
2. Si es la primera ejecución, abre Terminal en la raíz del proyecto y otorga permisos:

   ```bash
   chmod +x exec.sh
   ```
3. En el gestor de archivos (Finder, Nautilus, etc.), navega a la raíz y haz doble clic en `exec.sh`.
4. Se abrirá una ventana de terminal; responde cuando se solicite:

   * Nombre de la empresa.
   * Número de invitados.
5. Al finalizar, pulsa cualquier tecla para cerrar la ventana.
6. El informe generado se guardará en `Informes generados`.

## Configuración adicional

* **Plantilla de Word**:

  * No modifiques los nombres de los marcadores en `plantilla_burnout.docx` (por ejemplo `${EMPRESA}`, `${INTRINSECA}`, etc.).
* **Archivos JSON de parámetros** (`medidas.json`, `informacion_prl.json`):

  * Ajusta los rangos y valores por defecto según tus necesidades. Deben residir en la raíz del proyecto.

## Resolución de problemas

* **Error “conda no encontrado”**:

  * Verifica la instalación de Anaconda/Miniconda en la carpeta de usuario.
* **No se detectan archivos CSV**:

  * Asegúrate de que tengan extensión `.csv` y estén en `Respuestas`.
* **Problemas al crear o actualizar el entorno**:

  * Revisa la sintaxis de `environment.yml` y la configuración de conda.
