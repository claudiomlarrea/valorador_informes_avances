
# Cómo aplicar el FIX al Valorador de Informes de Avance (sin tocar la calculadora)

Este fix solo reemplaza **la exportación a Word** para evitar textos truncados con "..." al final del dictamen.

## Archivos
- `export_fix.py`  → Nueva función `export_word_dictamen(...)` sin recortes.
- `word_utils_fix.py` (opcional) → Utilidad `add_full_text` si querés usarla directo.
  
## Pasos (2 minutos, desde GitHub Web)
1) Subí `export_fix.py` (y `word_utils_fix.py` si querés) a la **raíz** del repo del valorador.
2) Abrí `app.py` en el editor web de GitHub.
3) Buscá la parte del botón de descarga del Word (ejemplo):
   ```python
   st.download_button(
       "Descargar informe Word",
       data=export_word(section_results, total_general, ...),
       file_name="informe_valoracion_avance.docx",
       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
   )
   ```
4) Reemplazá la función llamada dentro de `data=` por la nueva:
   ```python
   from export_fix import export_word_dictamen  # al inicio de app.py

   st.download_button(
       "Descargar informe Word",
       data=export_word_dictamen(section_results, total_general, dictamen_texto, categoria),
       file_name="informe_valoracion_avance.docx",
       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
   )
   ```
   - Asegurate de pasar el nombre real de tu variable de dictamen (p. ej. `dictamen_texto`, `texto_dictamen` o similar).  
   - Si tu app no usa `categoria`, podés pasar `""`.

5) Guardá cambios. En Streamlit Cloud hace **Rerun/Deploy** y listo.

> Nota: No tocamos nada de la **carga del informe**, **cálculo** ni **puntajes**. Solo cambiamos la forma de escribir el Word para que **no quede trunco**.

## Requisitos
`python-docx` ya debería estar en tus dependencias. No se necesita nada extra.
