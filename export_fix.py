# export_fix.py
# Versión final: agrega el nombre del proyecto en el Word sin alterar ningún cálculo.

from io import BytesIO
from datetime import datetime
from docx import Document
from docx.shared import Pt


def export_word(resultados, cumplimiento, dictamen_texto, categoria="", nombre_proyecto=""):
    """
    Genera el informe Word institucional con o sin nombre de proyecto.
    No modifica el cálculo del valorador.
    """
    doc = Document()
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Encabezado
    base_titulo = "UCCuyo – Valoración de Informe de Avance"
    if nombre_proyecto and str(nombre_proyecto).strip():
        titulo = f'{base_titulo} "Del proyecto {str(nombre_proyecto).strip()}"'
    else:
        titulo = base_titulo

    try:
        p = doc.add_paragraph(titulo)
        p.style = "Title"
        p.runs[0].font.size = Pt(14)
    except Exception:
        p = doc.add_paragraph(titulo)
        p.runs[0].font.size = Pt(14)

    doc.add_paragraph(f"Fecha: {fecha}")
    doc.add_paragraph("")  # Espacio

    # Resultados por criterio
    doc.add_paragraph("Resultados por criterio")
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Criterio"
    hdr_cells[1].text = "Puntaje (0–4)"

    for criterio, puntaje in resultados.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(criterio)
        row_cells[1].text = str(puntaje)

    # Cumplimiento y dictamen
    try:
        cumplimiento_txt = f"{float(cumplimiento):.1f}%"
    except Exception:
        cumplimiento_txt = str(cumplimiento)

    doc.add_paragraph(f"\nCumplimiento: {cumplimiento_txt}")
    doc.add_paragraph("\nDictamen final")
    if categoria:
        doc.add_paragraph(str(categoria))
    if dictamen_texto:
        doc.add_paragraph(str(dictamen_texto))

    doc.add_paragraph(
        "\nObservaciones del evaluador\n"
        + "..............................................................................\n"
        + "..............................................................................\n"
        + ".............................................................................."
    )

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
