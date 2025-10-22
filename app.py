
import streamlit as st
import re, json, io
import pandas as pd
from docx import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH

try:
    import pdfplumber
    HAVE_PDF = True
except Exception:
    HAVE_PDF = False

st.set_page_config(page_title="Valorador de Informes de Avance", layout="wide")
st.title("Valorador de Informes de Avance — UCCuyo")
st.caption("Carga .docx o .pdf, valora y exporta Excel/Word (sin recortar dictamen).")

@st.cache_data
def load_json_default(path: str, default: dict):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def extract_text_docx(file):
    doc = DocxDocument(file)
    text = "\\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += "\\n" + " | ".join(c.text for c in row.cells)
    return text

def extract_text_pdf(file):
    if not HAVE_PDF:
        raise RuntimeError("Para leer PDF instalá: pip install pdfplumber")
    chunks = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            chunks.append(p.extract_text() or "")
    return "\\n".join(chunks)

def match_count(pattern, text):
    return len(re.findall(pattern, text, re.IGNORECASE)) if pattern else 0

def clip(v, cap):
    return min(v, cap) if cap else v

# ---------- FIX Word (dictamen completo) ----------
def add_full_text(doc, text: str):
    if not text:
        return
    text = text.replace("\\r\\n", "\\n").replace("\\r", "\\n")
    blocks = text.split("\\n\\n")
    for block in blocks:
        for line in block.split("\\n"):
            doc.add_paragraph(line)
        doc.add_paragraph("")

def export_word_dictamen(section_results: dict, total_general: float, dictamen_texto: str, decision: str) -> bytes:
    from docx import Document
    doc = Document()
    p = doc.add_paragraph("Universidad Católica de Cuyo — Secretaría de Investigación")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Informe de valoración — Informe de Avance").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("")
    doc.add_paragraph(f"Resultado: {decision}")
    doc.add_paragraph(f"Puntaje total: {total_general:.1f}")
    doc.add_paragraph("")

    doc.add_heading("Dictamen", level=2)
    add_full_text(doc, dictamen_texto)

    for sec, data in section_results.items():
        doc.add_heading(sec, level=3)
        df = data.get("df", pd.DataFrame())
        if df.empty:
            doc.add_paragraph("Sin ítems detectados.")
        else:
            table = doc.add_table(rows=1, cols=len(df.columns))
            hdr = table.rows[0].cells
            for i, c in enumerate(df.columns):
                hdr[i].text = str(c)
            for _, row in df.iterrows():
                cells = table.add_row().cells
                for i, c in enumerate(df.columns):
                    cells[i].text = str(row[c])
        doc.add_paragraph(f"Subtotal sección: {data.get('subtotal', 0):.1f}")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ---------- Criterios por defecto (pueden sobreescribirse con criteria_avance.json) ----------
DEFAULT_CRITERIA = {
    "sections": {
        "Avances de cronograma: seguimiento/ajustes": {
            "max_points": 200,
            "items": {
                "Hitos cumplidos": {"pattern": r"(?i)hitos? cumplid", "unit_points": 10, "max_points": 100},
                "Cronograma ajustado": {"pattern": r"(?i)cronograma.*ajust", "unit_points": 10, "max_points": 100}
            }
        },
        "Producción parcial (borradores/ponencias)": {
            "max_points": 200,
            "items": {
                "Borradores/Manuscritos": {"pattern": r"(?i)(borrador|manuscrito)", "unit_points": 20, "max_points": 100},
                "Ponencias/Resúmenes": {"pattern": r"(?i)(ponenc|resumen ampliado)", "unit_points": 10, "max_points": 100}
            }
        },
        "Gestión y equipo [reuniones/tesistas]": {
            "max_points": 200,
            "items": {
                "Reuniones de equipo": {"pattern": r"(?i)reuniones? de equipo", "unit_points": 10, "max_points": 100},
                "Formación de tesistas/becarios": {"pattern": r"(?i)(tesistas?|becari)", "unit_points": 10, "max_points": 100}
            }
        },
        "Vinculación / Extensión": {
            "max_points": 100,
            "items": {
                "Actividades de extensión": {"pattern": r"(?i)extensi[oó]n|vinculaci[oó]n", "unit_points": 20, "max_points": 100}
            }
        },
        "Ejecución presupuestaria (uso de partidas)": {
            "max_points": 100,
            "items": {
                "Uso de partidas": {"pattern": r"(?i)(partidas?|rendici[oó]n)", "unit_points": 20, "max_points": 100}
            }
        }
    }
}
criteria = load_json_default("criteria_avance.json", DEFAULT_CRITERIA)

# ---------- Utilidad: nombre seguro de hoja de Excel ----------
INVALID_EXCEL_CHARS = r'[:\\\/\?\*\[\]]'
def safe_sheet_name(name: str, used: set) -> str:
    # limpiar inválidos y recortar
    cleaned = re.sub(INVALID_EXCEL_CHARS, " ", name).strip() or "Hoja"
    cleaned = cleaned[:31]
    base = cleaned
    i = 1
    while cleaned in used or cleaned == "RESUMEN":
        suffix = f"_{i}"
        cleaned = (base[:31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else (base + suffix)
        i += 1
    used.add(cleaned)
    return cleaned

# ===================== UI =====================
col1, col2 = st.columns([2,1])
with col1:
    file = st.file_uploader("Subí el Informe de Avance (.docx o .pdf)", type=["docx", "pdf"])
with col2:
    decision = st.selectbox("Resultado", ["Aprobado", "Aprobado con observaciones", "No aprobado"])

dictamen_texto = st.text_area("Dictamen (podés editarlo antes de exportar)", height=160, placeholder="Escribí aquí el dictamen final…")

if file:
    ext = (file.name.split(".")[-1] or "").lower()
    try:
        raw_text = extract_text_docx(file) if ext == "docx" else extract_text_pdf(file)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.success(f"Archivo cargado: {file.name}")
    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto del informe", raw_text, height=220)

    section_results = {}
    total_general = 0.0

    for section, cfg in criteria["sections"].items():
        st.markdown(f"### {section}")
        rows = []
        subtotal_items = 0.0
        for item, icfg in cfg.get("items", {}).items():
            c = match_count(icfg.get("pattern", ""), raw_text)
            pts = clip(c * icfg.get("unit_points", 0), icfg.get("max_points", 0))
            rows.append({"Ítem": item, "Ocurrencias": c, "Puntaje (tope ítem)": pts, "Tope ítem": icfg.get("max_points", 0)})
            subtotal_items += pts
        df = pd.DataFrame(rows)
        subtotal = clip(subtotal_items, cfg.get("max_points", 0))
        st.dataframe(df, use_container_width=True)
        st.info(f"Subtotal {section}: {subtotal} / máx {cfg.get('max_points', 0)}")
        section_results[section] = {"df": df, "subtotal": subtotal}
        total_general += subtotal

    st.markdown("---")
    st.subheader("Puntaje total")
    st.metric("Total acumulado", f"{total_general:.1f}")
    st.metric("Resultado", decision)

    st.markdown("---")
    st.subheader("Exportar resultados")

    # ---------- Excel con nombres de hoja seguros ----------
    out_xlsx = io.BytesIO()
    used_names = set()
    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as writer:
        for sec, data in section_results.items():
            sheet = safe_sheet_name(sec, used_names)
            data["df"].to_excel(writer, sheet_name=sheet, index=False)
        resumen = pd.DataFrame({
            "Sección": list(section_results.keys()),
            "Subtotal": [section_results[s]["subtotal"] for s in section_results]
        })
        resumen.loc[len(resumen)] = ["TOTAL", resumen["Subtotal"].sum()]
        resumen.loc[len(resumen)] = ["RESULTADO", decision]
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)

    st.download_button(
        "Descargar Excel",
        data=out_xlsx.getvalue(),
        file_name="valoracion_informe_avance.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.download_button(
        "Descargar informe Word",
        data=export_word_dictamen(section_results, total_general, dictamen_texto, decision),
        file_name="informe_valoracion_avance.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
else:
    st.info("Subí el Informe de Avance para valorar y exportar.")
