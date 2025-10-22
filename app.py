
import streamlit as st
import io, re, json, datetime
import pandas as pd
from docx import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
import yaml

# --------- PDF opcional ---------
try:
    import pdfplumber
    HAVE_PDF = True
except Exception:
    HAVE_PDF = False

st.set_page_config(page_title="Valorador de Informes de Avance", layout="wide")
st.title("Valorador de Informes de Avance — UCCuyo")
st.caption("Calcula 11 criterios ponderados y exporta Excel + Word (sin recortar texto).")

# --------- Config ---------
@st.cache_data
def load_yaml(path):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

cfg = load_yaml("rubric_config.yaml")
WEIGHTS = cfg.get("weights", {})
SCALE = cfg.get("scale", {"min": 0, "max": 4})
KEYWORDS = cfg.get("keywords", {})
TH = cfg.get("thresholds", {"aprobado": 70, "aprobado_obs": 50})

CRITERIA_ORDER = [
    "identificacion","cronograma","objetivos","metodologia","resultados",
    "formacion","gestion","dificultades","difusion","calidad_formal","impacto"
]

NAMES = {
    "identificacion": "Identificación general del proyecto",
    "cronograma": "Cumplimiento del cronograma",
    "objetivos": "Grado de cumplimiento de los objetivos",
    "metodologia": "Metodología",
    "resultados": "Resultados parciales",
    "formacion": "Formación de recursos humanos",
    "gestion": "Gestión del proyecto",
    "dificultades": "Dificultades y estrategias",
    "difusion": "Difusión y transferencia",
    "calidad_formal": "Calidad formal del informe",
    "impacto": "Impacto y proyección",
}

# --------- Utilidades extracción ---------
def extract_text_docx(file):
    doc = DocxDocument(file)
    text = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += "\n" + " | ".join(c.text for c in row.cells)
    return text

def extract_text_pdf(file):
    if not HAVE_PDF:
        raise RuntimeError("Para leer PDF: pip install pdfplumber")
    chunks = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            chunks.append(p.extract_text() or "")
    return "\n".join(chunks)

def auto_score(text, keywords, scale_max=4):
    text_low = text.lower()
    hits = 0
    for kw in keywords:
        if kw.lower() in text_low:
            hits += 1
    # mapa simple: 0 = 0; 1 = 2; 2+ = 3/4
    if hits == 0: return 0
    if hits == 1: return 2
    if hits == 2: return 3
    return min(scale_max, 4)

def decision_from(total_pct, th_ok, th_obs):
    if total_pct >= th_ok: return "APROBADO"
    if total_pct >= th_obs: return "APROBADO CON OBSERVACIONES"
    return "NO APROBADO"

INVALID_EXCEL_CHARS = r'[:\\\/\?\*\[\]]'
def safe_sheet_name(name: str, used: set) -> str:
    import re
    cleaned = re.sub(INVALID_EXCEL_CHARS, " ", name).strip() or "Hoja"
    cleaned = cleaned[:31]
    base = cleaned
    i = 1
    while cleaned in used or cleaned == "RESUMEN":
        suffix = f"_{i}"
        cleaned = (base[:31-len(suffix)] + suffix)
        i += 1
    used.add(cleaned)
    return cleaned

# --------- UI ---------
col1, col2 = st.columns([2,1])
with col1:
    up = st.file_uploader("Subí el Informe de Avance (.docx o .pdf)", type=["docx","pdf"])
with col2:
    st.write("Escala de cada criterio: 0–4")
    st.write("Pesos configurables en rubric_config.yaml")

dictamen_texto = st.text_area("Dictamen (podés editarlo antes de exportar)", height=160, placeholder="Escribí aquí el dictamen final…")

if up:
    ext = (up.name.split(".")[-1] or "").lower()
    try:
        text = extract_text_docx(up) if ext == "docx" else extract_text_pdf(up)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.success(f"Archivo cargado: {up.name}")
    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", text, height=200)

    # Autovaloración + ajustes manuales
    data = []
    total_pct = 0.0
    sliders = {}
    for key in CRITERIA_ORDER:
        section = NAMES[key]
        w = WEIGHTS.get(key, 0)
        kws = KEYWORDS.get(key, [])
        auto = auto_score(text, kws, SCALE.get("max",4))
        sliders[key] = st.slider(f"{section} (peso {w}%)", min_value=SCALE.get("min",0), max_value=SCALE.get("max",4),
                                 value=int(auto), step=1)
        contrib = sliders[key] / SCALE.get("max",4) * w
        total_pct += contrib
        data.append({"Criterio": section, "Puntaje (0-4)": sliders[key], "Peso (%)": w, "Aporte (%)": round(contrib,2)})

    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True)
    st.info(f"Cumplimiento: {round(total_pct,2)}%")

    dec = decision_from(total_pct, TH.get("aprobado",70), TH.get("aprobado_obs",50))
    st.metric("Dictamen", dec)

    # --------- Exportar Excel ---------
    out_xlsx = io.BytesIO()
    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as writer:
        used = set()
        sheet = safe_sheet_name("Valoración", used)
        df.to_excel(writer, sheet_name=sheet, index=False)
        resumen = pd.DataFrame({
            "Sección": df["Criterio"],
            "Puntaje (0-4)": df["Puntaje (0-4)"],
            "Peso (%)": df["Peso (%)"],
            "Aporte (%)": df["Aporte (%)"],
        })
        resumen.loc[len(resumen)] = ["TOTAL", "", "", round(resumen["Aporte (%)"].astype(float).sum(),2)]
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
    st.download_button("Descargar Excel", out_xlsx.getvalue(),
        file_name="valoracion_informe_avance.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

    # --------- Exportar Word (SIN recortar texto) ---------
    def add_full_text(doc, text: str):
        if not text: return
        text = text.replace("\r\n","\n").replace("\r","\n")
        blocks = text.split("\n\n")
        for block in blocks:
            for line in block.split("\n"):
                doc.add_paragraph(line)
            doc.add_paragraph("")

    def export_word():
        d = Document()
        p = d.add_paragraph("UCCuyo – Valoración de Informe de Avance")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        d.add_paragraph(f"Fecha: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        d.add_paragraph("")
        d.add_paragraph(f"Dictamen: {dec}  —  Cumplimiento: {round(total_pct,1)}%")
        d.add_paragraph("")
        d.add_paragraph("Resultados por criterio")
        for _, row in df.iterrows():
            d.add_paragraph(f"{row['Criterio']} (Puntaje: {row['Puntaje (0-4)']}/4 · Peso: {row['Peso (%)']}% · Aporte: {row['Aporte (%)']}%)")
        d.add_paragraph("")
        d.add_paragraph("Interpretación")
        # escribir el dictamen COMPLETO que el evaluador editó (sin 'shorten', sin '...')
        add_full_text(d, dictamen_texto)
        d.add_paragraph("")
        d.add_paragraph("Evidencia analizada (extracto)")
        # incluimos el texto completo o un extracto largo sin añadir '...'
        add_full_text(d, text)
        bio = io.BytesIO()
        d.save(bio)
        return bio.getvalue()

    st.download_button("Descargar Word", export_word(),
        file_name="dictamen_informe_avance.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)
else:
    st.info("Subí el informe para valorar y exportar.")
