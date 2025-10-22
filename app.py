
# app_dictamen_fix.py — miniapp para generar el Word sin cortes
import streamlit as st
from word_utils_fix import export_informe_avance

st.set_page_config(page_title="Exportar Dictamen (fix)", layout="wide")
st.title("Exportar Dictamen de Informe de Avance — Word sin cortes")

encabezado = st.text_input("Encabezado", "Universidad Católica de Cuyo — Secretaría de Investigación")
proyecto   = st.text_input("Proyecto (opcional)", "")
calif      = st.selectbox("Resultado", ["Aprobado", "Aprobado con observaciones", "No aprobado"]) 
dictamen   = st.text_area("Dictamen (pegar texto completo)", height=250)
interp     = st.text_area("Interpretación (opcional)", height=200)
obs        = st.text_area("Observaciones (opcional)", height=160)

if st.button("Generar Word", type="primary", use_container_width=True):
    if not dictamen.strip():
        st.error("Pegá al menos el dictamen.")
    else:
        path = "dictamen_informe_avance_COMPLETO.docx"
        export_informe_avance(path, encabezado, proyecto, calif, dictamen, interp, obs, tablas=None)
        with open(path, "rb") as f:
            st.download_button("Descargar Word", f.read(), file_name=path,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        st.success("Documento generado sin cortes.")
