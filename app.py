import io, re, yaml, math, pdfplumber
import streamlit as st
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Pt, Cm
from datetime import datetime

# =========================
# Configuración general
# =========================
st.set_page_config(
    page_title="UCCuyo · Valorador de Informes de Avance",
    page_icon="📊",
    layout="wide"
)

@st.cache_resource
def load_rubric():
    with open("rubric_config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

RUBRIC = load_rubric()
CRITERIA = [
    ("identificacion", "Identificación general del proyecto"),
    ("cronograma", "Cumplimiento del cronograma"),
    ("objetivos", "Grado de cumplimiento de los objetivos"),
    ("metodologia", "Metodología"),
    ("resultados", "Resultados parciales"),
    ("formacion", "Formación de recursos humanos"),
    ("gestion", "Gestión del proyecto"),
    ("dificultades", "Dificultades y estrategias"),
    ("difusion", "Difusión y transferencia"),
    ("calidad_formal", "Calidad formal del informe"),
    ("impacto", "Impacto y proyección"),
]

# =========================
# Funciones auxiliares
# =========================
def extract_text_from_docx(file_bytes: bytes) -> str:
    buffer = io.BytesIO(file_bytes)
    doc = Document(buffer)
    return "\n".join([p.text for p in doc.paragraphs])

def extract_text_from_pdf(file_bytes: bytes) -> str:
    buffer = io.BytesIO(file_bytes)
    text_parts = []
    with pdfplumber.open(buffer) as pdf:
        for page in pdf.pages:
            text_parts.append(page.extract_text() or "")
    return "\n".join(text_parts)

def naive_auto_score(text: str, key: str) -> int:
    """Heurística simple: cuenta coincidencias de palabras clave por criterio."""
    words = RUBRIC.get("keywords", {}).get(key, [])
    hits = 0
    lower = text.lower()
    for w in words:
        if w.lower() in lower:
            hits += 1
    if not words:
        return 0
    ratio = hits / len(words)
    if ratio == 0:
        return 0
    elif ratio < 0.25:
        return 1
    elif ratio < 0.5:
        return 2
    elif ratio < 0.75:
        return 3
    else:
        return 4

def weighted_total(scores: dict) -> float:
    weights = RUBRIC["weights"]
    total = 0.0
    for k, v in scores.items():
        w = weights.get(k, 0)
        total += (v / RUBRIC["scale"]["max"]) * w
    return round(total, 2)

def decision(final_pct: float) -> str:
    th = RUBRIC["thresholds"]
    if final_pct >= th["aprobado"]:
        return "APROBADO"
    elif final_pct >= th["aprobado_obs"]:
        return "APROBADO CON OBSERVACIONES"
    else:
        return "NO APROBADO"

def make_excel(scores: dict, final_pct: float, label: str) -> bytes:
    weights = RUBRIC["weights"]
    df = pd.DataFrame([{
        "Criterio": name,
        "Clave": key,
        "Puntaje (0-4)": scores[key],
        "Peso (%)": weights.get(key, 0),
        "Aporte (%)": round((scores[key]/RUBRIC["scale"]["max"])*weights.get(key,0), 2)
    } for key, name in CRITERIA])
    df_total = pd.DataFrame([{"Total (%)": final_pct, "Dictamen": label}])
    with io.BytesIO() as output:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Resultados")
            df_total.to_excel(writer, index=False, sheet_name="Resumen")
        return output.getvalue()

# =========================
# Funciones de texto
# =========================
def _add_full_text_as_paragraphs(doc: Document, text: str) -> None:
    """Agrega texto limpio, en párrafos legibles."""
    if not text:
        return
    blocks = re.split(r"\n{2,}", text.strip())
    for block in blocks:
        lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
        paragraph_text = " ".join(lines)
        if paragraph_text:
            p = doc.add_paragraph(paragraph_text)
            p.paragraph_format.space_after = Pt(6)
        else:
            doc.add_paragraph("")

# =========================
# Generación del Word
# =========================
def make_word(scores: dict, final_pct: float, label: str, raw_text: str) -> bytes:
    weights = RUBRIC["weights"]
    doc = Document()

    # --- Estilo base ---
    styles = doc.styles['Normal']
    styles.font.name = 'Times New Roman'
    styles.font.size = Pt(11)

    # --- Márgenes amplios ---
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    # --- Encabezado principal ---
    doc.add_heading('UCCuyo – Valoración de Informe de Avance', level=1)
    today = datetime.now().strftime("%Y-%m-%d %H:%M")
    doc.add_paragraph(f"Fecha: {today}")
    doc.add_paragraph(f"Dictamen: {label}  —  Cumplimiento: {final_pct}%")
    doc.add_paragraph("")
    doc.add_heading('Resultados por criterio', level=2)

    for key, name in CRITERIA:
        s = scores[key]
        w = weights.get(key, 0)
        aporte = round((s/RUBRIC['scale']['max'])*w, 2)
        p = doc.add_paragraph()
        run_title = p.add_run(f"{name} ")
        run_title.bold = True
        p.add_run(f"(Puntaje: {s}/4 · Peso: {w}% · Aporte: {aporte}%)")

    doc.add_paragraph("")
    doc.add_heading('Interpretación', level=2)
    fortalezas = [name for key, name in CRITERIA if scores[key] >= 3]
    mejoras = [name for key, name in CRITERIA if scores[key] <= 1]
    doc.add_paragraph("Fortalezas: " + (", ".join(fortalezas) if fortalezas else "no se identifican fortalezas destacadas."))
    doc.add_paragraph("Aspectos a mejorar: " + (", ".join(mejoras) if mejoras else "no se identifican aspectos críticos."))

    doc.add_paragraph("")
    doc.add_heading('Evidencia analizada (texto completo)', level=2)

    # --- Recorte corregido ---
    # Mantiene desde "INFORME DE AVANCE" hasta el primer separador fuerte (si existe)
    texto_fragmento = raw_text
    patron = "INFORME DE AVANCE"
    lower_raw = raw_text.lower()
    pos = lower_raw.find(patron.lower())
    if pos != -1:
        fragment = raw_text[pos:]
        # buscar posibles puntos de corte posteriores
        separadores = [
            "\n___", "Resultados parciales", "RESULTADOS PARCIALES",
            "\nII.-", "\nII .-", "\nII -", "\n\nII", "\n—", "\n- - -"
        ]
        cortes = [fragment.find(s) for s in separadores if fragment.find(s) != -1]
        stop = min(cortes) if cortes else -1
        texto_fragmento = fragment[:stop].strip() if stop != -1 else fragment.strip()

    # --- Agregar texto ---
    _add_full_text_as_paragraphs(doc, texto_fragmento)

    # --- Guardar en buffer ---
    with io.BytesIO() as buffer:
        doc.save(buffer)
        return buffer.getvalue()

# =========================
# Interfaz Streamlit
# =========================
st.markdown("## 📊 Valorador de Informes de Avance")
st.write("Subí un **PDF o DOCX** del informe de avance. La app extrae el texto, propone un puntaje automático por 11 criterios y te permite **ajustarlos manualmente** antes de exportar los resultados.")

uploaded = st.file_uploader("Cargar archivo (PDF o DOCX)", type=["pdf", "docx"])

raw_text = ""
if uploaded is not None:
    data = uploaded.read()
    if uploaded.name.lower().endswith(".docx"):
        raw_text = extract_text_from_docx(data)
    else:
        raw_text = extract_text_from_pdf(data)

    with st.expander("📄 Texto extraído (vista previa)"):
        st.text_area("Contenido", raw_text[:6000], height=280)

    st.divider()
    st.subheader("Evaluación automática + ajuste manual")

    cols = st.columns(3)
    auto_scores = {}
    for idx, (key, name) in enumerate(CRITERIA):
        if idx % 3 == 0:
            cols = st.columns(3)
        col = cols[idx % 3]
        with col:
            auto = naive_auto_score(raw_text, key)
            auto_scores[key] = auto

    st.write("**Sugerencia automática (0–4)**:", auto_scores)

    st.markdown("### Ajustar puntajes (0–4)")
    scores = {}
    for key, name in CRITERIA:
        scores[key] = st.slider(name, min_value=0, max_value=4, value=int(auto_scores.get(key,0)))

    final_pct = weighted_total(scores)
    label = decision(final_pct)
    st.markdown(f"### Resultado: **{label}** — Cumplimiento **{final_pct}%**")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("⬇️ Exportar Excel"):
            xls = make_excel(scores, final_pct, label)
            st.download_button("Descargar resultados.xlsx", data=xls, file_name="valoracion_informe_avance.xlsx")
    with c2:
        if st.button("⬇️ Exportar Word"):
            docx_bytes = make_word(scores, final_pct, label, raw_text)
            st.download_button("Descargar dictamen.docx", data=docx_bytes, file_name="dictamen_informe_avance.docx")
    with c3:
        st.download_button("Descargar configuración (YAML)", data=open("rubric_config.yaml","rb").read(), file_name="rubric_config.yaml")
else:
    st.info("Esperando archivo...")
