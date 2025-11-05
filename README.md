# Tablero Voluntariado

Aplicación Streamlit para explorar el dataset deduplicado y normalizado de voluntariado.

## Requisitos

- Python 3.9+ (Cloud usa una versión reciente)
- Dependencias en `requirements.txt`:
  - `streamlit`, `pandas`, `numpy`, `openpyxl`, `pyarrow`

## Archivos de datos

- Archivo principal: `Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx`
- Hoja: `Merged`
- La app detecta y usa `País (normalizado)` si está presente.

## Ejecutar localmente

- Crear/activar entorno: `python -m venv .venv && source .venv/bin/activate`
- Instalar deps: `pip install -r requirements.txt`
- Ejecutar: `streamlit run streamlit_app.py`
- URL local: `http://localhost:8501`

## Publicar en Streamlit Cloud

1. Sube este repositorio a GitHub (incluyendo el Excel normalizado).
2. En Streamlit Cloud, crea una nueva app seleccionando el repo y el archivo principal `streamlit_app.py`.
3. Asegúrate de que `requirements.txt` esté presente (Cloud instalará automáticamente).
4. Opcional: en `.streamlit/config.toml` ya se fuerza el tema claro:
   - `[theme]\nbase = "light"`
5. Si el Excel es pesado, considera alojarlo en un bucket y leerlo vía URL (no requerido por ahora).

## Notas técnicas

- Cálculo de edad: parsea fechas en formato día/mes/año (`dayfirst=True`).
- Compatibilidad Arrow: las columnas tipo `object` se convierten a `str` al renderizar la tabla.
- Deprecación de ancho: `st.dataframe(..., width='stretch')` reemplaza `use_container_width`.

## Personalizaciones rápidas

- Top países: control en el sidebar para elegir N.
- Histograma de edades: control de bins y toggle para mostrar "Sin edad".
- Áreas de interés: top 10 del filtrado actual.