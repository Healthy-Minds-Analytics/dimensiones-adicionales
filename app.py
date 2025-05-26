# app.py
import streamlit as st
from Generar_informe_Satisfaccion import generar_informe_satisfaccion
from Generar_informe_Burnout import generar_informe_burnout

st.set_page_config(page_title="Generador de Informes", layout="wide")

st.title("📝 Generador de Informes Word")

# 1) Selección de informe
report_type = st.selectbox(
    "¿Qué informe quieres generar?",
    ["Satisfacción laboral", "Burnout"]
)

st.markdown("---")

# 2) Campos comunes (si los hubiera)
csv_file = st.file_uploader("Sube tu archivo CSV", type="csv")
empresa = st.text_input("Nombre de la empresa")
invitados = st.number_input("Número de invitados", min_value=1, value=1)

# 3) Campos específicos según informe
if report_type == "Satisfacción laboral":
    st.subheader("Parámetros – Satisfacción laboral")
    # Por ejemplo:
    num_medidas = st.slider("Número de medidas a proponer", 1, 10, 3)
    # Si tiene JSON de agrupación fijo, no pides; si es variable:
    # json_agrup = st.file_uploader("JSON de criterios", type="json")
elif report_type == "Burnout":
    st.subheader("Parámetros – Burnout")
    limite_alerta = st.slider("Límite para alertas", 1, 15, 10)
    # y más inputs que requiera Burnout…

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
        else:  # Burnout
            docx_bytes = generar_informe_burnout(
                csv_file,
                empresa=empresa,
                invitados=invitados,
                limite=limite_alerta,
                # …otros params…
            )
            filename = f"Burnout_{empresa}.docx"

        # 5) Descarga directa
        st.download_button(
            label="📥 Descargar informe Word",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
