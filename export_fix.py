
# export_fix.py — Exportación a Word sin truncado para Informes de Avance
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import pandas as pd

def add_full_text(doc, text: str):
    """Escribe 'text' completo preservando párrafos. No acorta ni añade '...'."""
    if not text:
        return
    text = text.replace('\r\n','\n').replace('\r','\n')
    blocks = text.split('\n\n')
    for block in blocks:
        for line in block.split('\n'):
            doc.add_paragraph(line)
        doc.add_paragraph('')  # separación entre bloques

def export_word_dictamen(section_results: dict, total_general: float, dictamen_texto: str, categoria: str) -> bytes:
    """Genera el Word final del dictamen sin recortes.
    - section_results: {seccion: {'df': pandas.DataFrame, 'subtotal': float}, ...}
    - total_general: puntaje total
    - dictamen_texto: texto completo del dictamen (se escribirá íntegro)
    - categoria: categoría alcanzada (si aplica)
    Devuelve: bytes del .docx listo para 'st.download_button'.
    """
    doc = Document()
    p = doc.add_paragraph('Universidad Católica de Cuyo — Secretaría de Investigación')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('Informe de valoración — Informe de Avance').alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('')
    doc.add_paragraph(f'Puntaje total: {total_general:.1f}')
    if categoria:
        doc.add_paragraph(f'Categoría alcanzada: {categoria}')
    doc.add_paragraph('')

    # Dictamen completo (sin cortar)
    doc.add_heading('Dictamen', level=2)
    add_full_text(doc, dictamen_texto)

    # Tablas por sección
    for sec, data in section_results.items():
        doc.add_heading(sec, level=3)
        df = data.get('df')
        if df is None or df.empty:
            doc.add_paragraph('Sin ítems detectados.')
        else:
            table = doc.add_table(rows=1, cols=len(df.columns))
            hdr = table.rows[0].cells
            for i, c in enumerate(df.columns):
                hdr[i].text = str(c)
            for _, row in df.iterrows():
                cells = table.add_row().cells
                for i, c in enumerate(df.columns):
                    cells[i].text = str(row[c])
        doc.add_paragraph(f'Subtotal sección: {data.get("subtotal", 0):.1f}')

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
