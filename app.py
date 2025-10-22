
import streamlit as st
import io, re, json, datetime
import pandas as pd
from docx import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx import Document
import yaml

# --------- PDF opcional (no afecta puntajes) ---------
try:
    import pdfplumber
    HAVE_PDF = True
except Exception:
    HAVE_PDF = False

st.set_page_config(page_title="Valorador de Informes de Avance", layout="wide")
st.title("Valorador de Informes de Avance — UCCuyo")
st.caption("Se conservan rúbrica y umbrales del proyecto. Exportación Word sin truncados, con sangría y justificado.")

# --------- Config ---------
@st.cache_data
def load_yaml(path):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

cfg = load_yaml("rubric_config.yaml")
SCALE = cfg.get("scale", {"min": 0, "max": 4})
WEIGHTS = cfg.get("weights", {})
TH = cfg.get("thresholds", {"aprobado": 70, "aprobado_obs": 50})
KEYWORDS = cfg.get("keywords", {})  # solo informativo (no altera puntaje)

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
    text = "\\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += "\\n" + " | ".join(c.text for c in row.cells)
    return text

def extract_text_pdf(file):
    if not HAVE_PDF:
        raise RuntimeError("Para leer PDF: pip install pdfplumber")
    chunks = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            chunks.append(p.extract_text() or "")
    return "\\n".join(chunks)

def decision_from(total_pct, th_ok, th_obs):
    if total_pct >= th_ok: return "APROBADO"
    if total_pct >= th_obs: return "APROBADO CON OBSERVACIONES"
    return "NO APROBADO"

INVALID_EXCEL_CHARS = r'[:\\\/\?\*\[\]]'
def safe_sheet_name(name: str, used: set) -> str:
    cleaned = re.sub(INVALID_EXCEL_CHARS, " ", name).strip() or "Hoja"
    cleaned = cleaned[:31]
    base = cleaned
    i = 1
    while cleaned in used or cleaned == "RESUMEN":
        suffix = f"_{i}"
        cleaned = (base[:31-len(suffix)] + suffix)
        i += 1
    used.add(cleaned); return cleaned

# --------- UI ---------
col1, col2 = st.columns([2,1])
with col1:
    up = st.file_uploader("Informe de Avance (.docx o .pdf)", type=["docx","pdf"])
with col2:
    st.write("Escala 0–4 (manual). Pesos en rubric_config.yaml")

dictamen_texto = st.text_area("Dictamen (se respeta formato en Word)", height=180, placeholder="Pegá o escribí el dictamen aquí…")

if up:
    ext = (up.name.split(".")[-1] or "").lower()
    try:
        texto_informe = extract_text_docx(up) if ext == "docx" else extract_text_pdf(up)
    except Exception as e:
        st.error(str(e)); st.stop()

    with st.expander("Ver texto extraído (solo referencia)"):
        st.text_area("Texto", texto_informe, height=220)

    # --------- Valoración (manual, SIN auto-asignación) ---------
    data = []
    total_pct = 0.0
    for key in CRITERIA_ORDER:
        section = NAMES[key]
        w = WEIGHTS.get(key, 0)
        # sugerencia (no vinculante) por cantidad de keywords encontradas
        sugerencia = sum(1 for kw in KEYWORDS.get(key, []) if kw.lower() in texto_informe.lower())
        help_txt = f"Sugerencia orientativa por palabras clave: {min(sugerencia, SCALE.get('max',4))}/4 (no afecta el puntaje)."
        val = st.slider(f"{section} (peso {w}%)", min_value=SCALE.get("min",0), max_value=SCALE.get("max",4), value=0, step=1, help=help_txt)
        aporte = (val / SCALE.get("max",4)) * w if w else 0.0
        total_pct += aporte
        data.append({"Criterio": section, "Puntaje (0-4)": val, "Peso (%)": w, "Aporte (%)": round(aporte,2)})
    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True)

    st.subheader("Resultado")
    st.metric("Cumplimiento", f"{round(total_pct,1)}%")
    dec = decision_from(total_pct, TH.get("aprobado",70), TH.get("aprobado_obs",50))
    st.metric("Dictamen", dec)

    # --------- Exportar Excel ---------
    out_xlsx = io.BytesIO()
    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as writer:
        used = set()
        df.to_excel(writer, sheet_name=safe_sheet_name("Valoración", used), index=False)
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

    # --------- Exportar Word (sin truncado + sangría y justificado) ---------
    def add_blocks_as_paragraphs(doc, text, indent_cm=0.75):
        if not text: return
        text = text.replace("\r\n","\n").replace("\r","\n")
        blocks = [b.strip() for b in text.split("\n\n") if b.strip()]
        for block in blocks:
            # unimos líneas internas del pegado para no cortar párrafos
            paragraph = block.replace("\n", " ").strip()
            p = doc.add_paragraph(paragraph)
            p.paragraph_format.first_line_indent = Cm(indent_cm)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def export_word():
        d = Document()
        header = d.add_paragraph("UCCuyo – Valoración de Informe de Avance")
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        d.add_paragraph(f"Fecha: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        d.add_paragraph("")
        d.add_paragraph(f"Dictamen: {dec}  —  Cumplimiento: {round(total_pct,1)}%")
        d.add_paragraph("")

        d.add_paragraph("Resultados por criterio")
        for _, row in df.iterrows():
            p = d.add_paragraph(f"{row['Criterio']} (Puntaje: {row['Puntaje (0-4)']}/4 · Peso: {row['Peso (%)']}% · Aporte: {row['Aporte (%)']}%)")
            p.paragraph_format.first_line_indent = Cm(0.75)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        d.add_paragraph("")
        d.add_paragraph("Interpretación")
        add_blocks_as_paragraphs(d, dictamen_texto, indent_cm=0.75)

        d.add_paragraph("")
        d.add_paragraph("Evidencia analizada (extracto)")
        add_blocks_as_paragraphs(d, texto_informe, indent_cm=0.75)

        bio = io.BytesIO(); d.save(bio); return bio.getvalue()

    st.download_button("Descargar Word", export_word(),
                       file_name="dictamen_informe_avance.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True)
else:
    st.info("Subí el informe para valorar y exportar.")
