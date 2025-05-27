# app.py
import streamlit as st
import json
from Generar_informe_Satisfaccion import generar_informe_satisfaccion
from Generar_informe_Burnout import generar_informe_burnout
from Generar_informe_Generico import generar_informe_generico

st.set_page_config(page_title="Generador de Informes", layout="wide")

st.title("📝 Generador de Informes Word")

# 1) Selección de informe
report_type = st.selectbox(
    "¿Qué informe quieres generar?",
    ["Satisfacción laboral", "Burnout", "Genérico"]
)

st.markdown("---")

# 2) Campos comunes
csv_file = st.file_uploader("Sube tu archivo CSV", type="csv")
empresa = st.text_input("Nombre de la empresa")
invitados = st.number_input("Número de invitados", min_value=1, value=1)

# 3) Campos específicos según informe
if report_type == "Satisfacción laboral":
    st.subheader("Parámetros – Satisfacción laboral")
    num_medidas = st.slider("Número de medidas a proponer", 1, 10, 3)
elif report_type == "Burnout":
    st.subheader("Parámetros – Burnout")
    limite_alerta = st.slider("Límite para alertas", 1, 15, 10)
    # y más inputs que requiera Burnout…
elif report_type == "Genérico":
    titulo = st.text_input("Título del informe")
    json_file = st.file_uploader("JSON de preguntas", type="json")
    if json_file:
        # Carga en memoria
        json_data = json.load(json_file)
        # Extrae la lista de locales
        locales = json_data.get("availableLocales", [])
        # Desplegable con esos valores
        locale = st.selectbox("Idioma del informe", locales)
    else:
        st.info("Sube el JSON de preguntas para elegir idioma")

st.markdown("---")

# 4) Botón de generación
if st.button("▶️ Generar informe"):
    # Validaciones básicas
    if not csv_file:
        st.error("❌ Debes subir primero un archivo CSV.")
    else:
        # Llamada exclusiva según la elección
        if report_type == "Satisfacción laboral":
            docx_bytes = generar_informe_satisfaccion(
                csv_file,
                empresa=empresa,
                invitados=invitados,
                num_medidas=num_medidas,
                # …otros params…
            )
            filename = f"Satisfaccion_{empresa}.docx"
        elif report_type == "Burnout":  # Burnout
            docx_bytes = generar_informe_burnout(
                csv_file,
                empresa=empresa,
                invitados=invitados,
                limite=limite_alerta,
            )
            filename = f"Burnout_{empresa}.docx"

        else: #if report_type == "Generico":  # Genérico
            if not json_file:
                st.error("❌ Debes subir un JSON de preguntas")
            elif not titulo:
                st.error("❌ Debes ingresar un título para el informe")
            elif not empresa:
                st.error("❌ Debes ingresar el nombre de la empresa")
            elif not locale:
                st.error("❌ Debes elegir un idioma para el informe")
            else:
                docx_bytes = generar_informe_generico(
                    csv_source=csv_file,
                    json_source=json_file,
                    empresa=empresa,
                    titulo=titulo,
                    invitados=invitados,
                    locale=locale,
                )
                filename = f"{titulo.replace(' ','_')}_{empresa}.docx"


        # 5) Descarga directa
        st.download_button(
            label="📥 Descargar informe Word",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
