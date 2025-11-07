
# app.py — Valorador de Informes de Avance (versión con nombre de proyecto en el Word)
import io
import re
import yaml
import pdfplumber
import streamlit as st
from docx import Document as DocxDocument

from export_fix import export_word_dictamen


# ---------- Utilidades ----------
def load_config(path: str = "rubric_config.yaml") -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def extract_text(uploaded_file) -> str:
    """Extrae texto básico de PDF o DOCX."""
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        try:
            text = []
            with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
                for page in pdf.pages:
                    text.append(page.extract_text() or "")
            return "\n".join(text)
        except Exception:
            return ""
    elif name.endswith(".docx"):
        try:
            doc = DocxDocument(io.BytesIO(uploaded_file.read()))
            return "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            return ""
    else:
        # Otros formatos: intentar leer como texto
        try:
            return uploaded_file.read().decode("utf-8", errors="ignore")
        except Exception:
            return ""


def auto_score(text: str, keywords_map: dict, scale_min: int, scale_max: int) -> dict:
    """Scoring simple por presencia de palabras clave (AND/OR básico)."""
    text_norm = text.lower()
    scores = {}
    for key, kw_list in keywords_map.items():
        hits = 0
        for token in kw_list:
            if re.search(r"\b" + re.escape(token.lower()) + r"\b", text_norm):
                hits += 1
        # Normalizar a la escala 0–max
        ratio = min(hits / len(kw_list), 1.0) if kw_list else 0.0
        score = round(scale_min + ratio * (scale_max - scale_min))
        scores[key] = int(score)
    return scores


def pretty_names():
    return {
        "identificacion": "Identificacion",
        "cronograma": "Cronograma",
        "objetivos": "Objetivos",
        "metodologia": "Metodologia",
        "resultados": "Resultados",
        "formacion": "Formacion",
        "gestion": "Gestion",
        "dificultades": "Dificultades",
        "difusion": "Difusion",
        "calidad_formal": "Calidad formal",
        "impacto": "Impacto",
    }


def compute_cumplimiento(puntajes: dict, cfg: dict) -> float:
    scale_max = cfg["scale"]["max"]
    weights = cfg["weights"]
    s = 0.0
    for k, w in weights.items():
        score = puntajes.get(k, 0)
        s += (score / scale_max) * w
    return float(s)  # por construcción, 0–100


def dictamen_from_cumplimiento(p: float, cfg: dict) -> str:
    th_ok = cfg["thresholds"]["aprobado"]
    th_obs = cfg["thresholds"]["aprobado_obs"]
    if p >= th_ok:
        return "Aprobado"
    if p >= th_obs:
        return "Aprobado con observaciones"
    return "No aprobado"


# ---------- App ----------
st.set_page_config(page_title="Valorador de Informes de Avance", layout="centered")
st.title("Valorador de Informes de Avance")

cfg = load_config("rubric_config.yaml")
scale_min = cfg["scale"]["min"]
scale_max = cfg["scale"]["max"]
weights = cfg["weights"]
keywords_map = cfg.get("keywords", {})

uploaded = st.file_uploader("Cargar informe (PDF o DOCX)", type=["pdf", "docx"])

# Campo para el nombre del proyecto (se usará en el DOCX)
nombre_proyecto = st.text_input(
    "Nombre del proyecto de investigación valorado (aparecerá en el Word):",
    value=st.session_state.get("nombre_proyecto", ""),
)
st.session_state["nombre_proyecto"] = nombre_proyecto

if uploaded:
    raw_text = extract_text(uploaded)
    if not raw_text.strip():
        st.warning("No se pudo extraer texto. Podés continuar ajustando manualmente los puntajes.")
    # Puntaje automático preliminar
    auto = auto_score(raw_text, keywords_map, scale_min, scale_max)

    st.subheader("Ajuste de puntajes por criterio (0–4)")
    pretty = pretty_names()
    puntajes = {}
    for key in weights.keys():  # mantener orden de la rúbrica
        default = auto.get(key, 0)
        puntajes[key] = st.slider(pretty[key], min_value=scale_min, max_value=scale_max, value=int(default))

    # Cálculo de cumplimiento y dictamen
    cumplimiento = compute_cumplimiento(puntajes, cfg)
    dictamen = dictamen_from_cumplimiento(cumplimiento, cfg)

    # Mostrar tabla de resultados
    st.subheader("Resultados por criterio")
    st.write("| Criterio | Puntaje (0–4) |")
    st.write("|---|---|")
    for key in weights.keys():
        st.write(f"| {pretty[key]} | {puntajes[key]} |")

    st.write(f"**Cumplimiento:** {cumplimiento:.1f}%")
    st.write("**Dictamen final:** ", dictamen)

    # Preparar dict para Word con nombres legibles (en el mismo orden)
    resultados_word = {pretty[k]: puntajes[k] for k in weights.keys()}

    # Botón de exportación a Word (usa export_word_dictamen e incluye el nombre del proyecto)
    st.download_button(
        "Descargar informe Word",
        data=export_word_dictamen(
            resultados=resultados_word,
            cumplimiento=cumplimiento,
            dictamen_texto=dictamen,
            categoria="",
            nombre_proyecto=nombre_proyecto,
        ),
        file_name="informe_valoracion_avance.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
else:
    st.info("Cargá un informe para comenzar. Luego ajustá los puntajes y descargá el Word.")
