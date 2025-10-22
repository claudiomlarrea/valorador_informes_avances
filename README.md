# Valorador de Informes de Avance- UCCuyo

Calculadora institucional para **valorar informes de avance** de proyectos de investigación a partir de un **PDF o Word**.
- Carga el archivo
- Extrae el texto automáticamente
- Evalúa 11 criterios (0–4) con **ponderaciones configurables** (`rubric_config.yaml`)
- Permite **ajuste manual** de puntajes por el evaluador
- Genera **Excel** con los puntajes y **Word** interpretativo con dictamen (Aprobado / Aprobado con observaciones / No aprobado)

## Criterios
1. Identificación general del proyecto  
2. Cumplimiento del cronograma  
3. Grado de cumplimiento de los objetivos  
4. Metodología  
5. Resultados parciales  
6. Formación de recursos humanos  
7. Gestión del proyecto  
# Fix de exportación a Word — Valorador de Informes de Avance

Este parche evita que los dictámenes queden truncados con puntos suspensivos (...).

## Archivos
- `word_utils_fix.py`: utilidades seguras para escribir texto completo al DOCX.
- `app_dictamen_fix.py`: miniapp Streamlit para generar el Word sin cortes.

## Uso rápido (miniapp)
```bash
pip install streamlit python-docx
streamlit run app_dictamen_fix.py
```

## Integración en tu app principal
1. Copiá `word_utils_fix.py` a tu proyecto.
2. En tu exportador, reemplazá la escritura del dictamen por:

```python
from word_utils_fix import add_full_text

# ...
doc.add_paragraph("Dictamen")
add_full_text(doc, texto_del_dictamen)   # sin shorten ni [:N] + '...'
# ...
```

3. Si generás tablas, usá `add_table(doc, headers, rows)`.

Con esto, el Word final no corta el texto ni añade puntos suspensivos.

8. Dificultades y estrategias  
9. Difusión y transferencia  
10. Calidad formal del informe  
11. Impacto y proyección  

## Ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura
- `app.py` — Aplicación Streamlit
- `rubric_config.yaml` — Pesos/umbral y palabras clave
- `requirements.txt` — Dependencias
- `runtime.txt` — Versión de Python para Streamlit Cloud
