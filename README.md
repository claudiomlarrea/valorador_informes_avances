# Valorador de Informes de Avance (Streamlit)

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
