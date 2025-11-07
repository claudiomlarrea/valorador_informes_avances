from datetime import datetime
from docx import Document
from docx.shared import Pt


def export_word_dictamen(resultados, cumplimiento, dictamen_texto, categoria="", nombre_proyecto=""):
    """
    Genera el informe Word institucional con el encabezado:
    'UCCuyo – Valoración de Informe de Avance "Del proyecto …"'
    sin alterar la lógica de cálculo del valorador.
    """
    doc = Document()

    # ---- ENCABEZADO INSTITUCIONAL ----
    titulo = f'UCCuyo – Valoración de Informe de Avance "Del proyecto {nombre_proyecto}"'
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")

    p = doc.add_paragraph(titulo)
    p.style = "Title"
    p.runs[0].font.size = Pt(14)

    doc.add_paragraph(f"Fecha: {fecha}")
    doc.add_paragraph("")  # espacio visual

    # ---- TABLA DE RESULTADOS ----
    doc.add_paragraph("Resultados por criterio")
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Criterio"
    hdr_cells[1].text = "Puntaje (0–4)"

    for criterio, puntaje in resultados.items():
        row_cells = table.add_row().cells
        row_cells[0].text = criterio.capitalize()
        row_cells[1].text = str(puntaje)

    # ---- CUMPLIMIENTO Y DICTAMEN ----
    doc.add_paragraph(f"\nCumplimiento: {cumplimiento:.1f}%")
    doc.add_paragraph("\nDictamen final")
    doc.add_paragraph(dictamen_texto)

    # ---- OBSERVACIONES ----
    doc.add_paragraph(
        "\nObservaciones del evaluador\n"
        + "..............................................................................\n"
        + "..............................................................................\n"
        + ".............................................................................."
    )

    return doc
