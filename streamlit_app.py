import datetime
import numpy as np
import os
import re
from io import BytesIO
from typing import Optional, Tuple, List

import pandas as pd
from PIL import Image
import streamlit as st
import altair as alt


# Configuración de página con logo si existe
logo_path = "assets/logo.png"
page_icon = None

if os.path.exists(logo_path):
    try:
        page_icon = Image.open(logo_path)
    except Exception:
        pass

if page_icon:
    st.set_page_config(page_title="Tablero Voluntariado", page_icon=page_icon, layout="wide")
else:
    st.set_page_config(page_title="Tablero Voluntariado", layout="wide")

st.title("Tablero de Voluntariado")

st.markdown("""
<style>
    /* Metric Cards */
    div.css-1r6slb0.e1tzin5v2 {
        background-color: #ffffff;
        border: 1px solid #dedede;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    
    .metric-card {
        background-color: #FFFFFF;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #E0E0E0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.04);
        text-align: center;
        margin-bottom: 10px;
    }
    .metric-label {
        color: #6b7280;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        font-weight: 600;
        margin-bottom: 5px;
    }
    .metric-value {
        color: #0F9D58;
        font-size: 2rem;
        font-weight: 700;
        margin: 0;
    }
</style>
""", unsafe_allow_html=True)


def _master_file_version(path: str) -> Optional[Tuple[float, int]]:
    """Devuelve (mtime, size) del archivo para usar como clave de caché. Si no existe, None."""
    try:
        st = os.stat(path)
        return (st.st_mtime, st.st_size)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def load_data(file_version: Optional[Tuple[float, int]]) -> Tuple[pd.DataFrame, Optional[str]]:
    """Carga el maestro. file_version (mtime, size) invalida la caché cuando el archivo cambia."""
    path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
    try:
        # Intentar leer 'Merged' o la primera hoja
        try:
            df = pd.read_excel(path, sheet_name="Merged")
        except ValueError:
            df = pd.read_excel(path, sheet_name=0)
        return df, path
    except Exception:
        return pd.DataFrame(), None


def ensure_fullname(df: pd.DataFrame) -> pd.DataFrame:
    first_col = "Nombre completo: First"
    last_col = "Nombre completo: Last"
    if first_col in df.columns or last_col in df.columns:
        first = df.get(first_col, "").fillna("").astype(str).str.strip()
        last = df.get(last_col, "").fillna("").astype(str).str.strip()
        df["Nombre completo"] = (first + " " + last).str.strip()
    elif "Nombre completo" not in df.columns:
        df["Nombre completo"] = df.index.astype(str)
    return df


def compute_age(d: pd.Timestamp) -> Optional[int]:
    try:
        if pd.isna(d):
            return None
        today = datetime.date.today()
        # Interpretar fechas en formato día/mes/año para evitar warnings
        dob_ts = pd.to_datetime(d, dayfirst=True, errors="coerce")
        if pd.isna(dob_ts):
            return None
        dob = dob_ts.date()
        return today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    except Exception:
        return None


def filter_by_areas(df: pd.DataFrame, selected_areas: List[str], match_all: bool) -> pd.Series:
    if not selected_areas:
        return pd.Series([True] * len(df))
    col = "Áreas de interés (lista)"
    if col not in df.columns:
        return pd.Series([True] * len(df))
    lists = df[col].fillna("").astype(str)
    def has_areas(s: str) -> bool:
        items = [x.strip() for x in s.split(";") if x.strip()]
        if match_all:
            return all(a in items for a in selected_areas)
        return any(a in items for a in selected_areas)
    return lists.apply(has_areas)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Merged") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    bio.seek(0)
    return bio.read()


# Carga de datos
_master_path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
_master_version = _master_file_version(_master_path)
df, source_path = load_data(_master_version)
if source_path is None or df.empty:
    st.error("No se encontró el archivo maestro: 'Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx'.")
    st.stop()

df = ensure_fullname(df.copy())


def ensure_arrow_compatible(df_in: pd.DataFrame) -> pd.DataFrame:
    df_out = df_in.copy()
    # Convertir todas las columnas de tipo object a string para evitar errores de Arrow
    for c in df_out.columns:
        if df_out[c].dtype == object:
            df_out[c] = df_out[c].astype(str)
    return df_out


# Sidebar: controles y filtros
if os.path.exists(logo_path):
    st.sidebar.image(logo_path, width="stretch")

st.sidebar.header("Controles")

# Búsqueda por nombre
query = st.sidebar.text_input("Buscar por nombre")

# Búsqueda por identificación
# Detectar columna de identificación de forma flexible
id_candidates = [
    "Identificación",
    "Identificacion",
    "Número de identificación",
    "Numero de identificacion",
    "Documento",
    "DNI",
    "Cédula",
    "Cedula",
    "ID",
    "Unnamed: 8",
]
id_col = next((c for c in id_candidates if c in df.columns), None)
id_query = st.sidebar.text_input("Buscar por identificación")
if id_query and not id_col:
    st.sidebar.caption("Columna de identificación no encontrada en el dataset")

# Filtros categóricos
sexo_vals = sorted([x for x in df.get("Sexo", pd.Series()).dropna().astype(str).unique()])
sexo_sel = st.sidebar.multiselect("Sexo", options=sexo_vals, default=[])

nivel_vals = sorted([x for x in df.get("Nivel académico", pd.Series()).dropna().astype(str).unique()])
nivel_sel = st.sidebar.multiselect("Nivel académico", options=nivel_vals, default=[])

# País (detección robusta de columna, preferir normalizado)
pais_candidates = [
    "País (normalizado)",
    "País",
    "Pais",
    "País de residencia",
    "País (Residencia)",
    "Country",
]
pais_col = next((c for c in pais_candidates if c in df.columns), None)
if pais_col:
    pais_vals = sorted([x for x in df.get(pais_col, pd.Series()).dropna().astype(str).str.strip().unique() if x])
    label = "País (normalizado)" if pais_col == "País (normalizado)" else "País"
    pais_sel = st.sidebar.multiselect(label, options=pais_vals, default=[])
else:
    st.sidebar.write("País: columna no encontrada")
    pais_sel = []

# Idiomas: detectar columnas por patrón y extraer idioma y nivel
lang_pattern = re.compile(r"Idiomas.*:\s*(.+?)\s+(B[áa]sico|Intermedio|Avanzado)", re.IGNORECASE)

def strip_accents(s: str) -> str:
    import unicodedata
    return (
        unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    )

def canon_lang(name: str) -> str:
    n = strip_accents(str(name).strip().lower())
    mapping = {
        "ingles": "Inglés",
        "espanol": "Español",
        "frances": "Francés",
        "portugues": "Portugués",
    }
    return mapping.get(n, name.strip().title())

def canon_level(level: str) -> str:
    l = strip_accents(str(level).strip().lower())
    if l.startswith("basico"):
        return "Básico"
    if l.startswith("intermedio"):
        return "Intermedio"
    if l.startswith("avanzado"):
        return "Avanzado"
    return level.strip().title()

lang_cols = []  # tuples: (col_name, language, level)
for c in df.columns:
    m = lang_pattern.search(str(c))
    if m:
        lang_raw, level_raw = m.group(1), m.group(2)
        lang_cols.append((c, canon_lang(lang_raw), canon_level(level_raw)))

languages = sorted(list({lc[1] for lc in lang_cols}))
levels = ["Básico", "Intermedio", "Avanzado"]

if languages:
    lang_sel = st.sidebar.multiselect("Idiomas", options=languages, default=[])
    level_sel = st.sidebar.multiselect("Nivel de idioma", options=levels, default=[])
else:
    lang_sel = []
    level_sel = []

# Áreas de interés
areas_col = "Áreas de interés (lista)"
areas_options: List[str] = []
if areas_col in df.columns:
    all_areas = (
        df[areas_col]
        .dropna()
        .astype(str)
        .str.split(";")
        .explode()
        .str.strip()
    )
    areas_options = sorted([x for x in all_areas.unique() if x])
areas_sel = st.sidebar.multiselect("Áreas de interés", options=areas_options, default=[])
match_all = st.sidebar.checkbox("Coincidir todas las áreas seleccionadas", value=False)

# Edad
edad_col = "Edad (calculada)"
if "Fecha de nacimiento" in df.columns:
    df[edad_col] = df["Fecha de nacimiento"].apply(compute_age)
else:
    df[edad_col] = None
min_age = 0
max_age = 100
age_range = st.sidebar.slider("Rango de edad", min_value=min_age, max_value=max_age, value=(min_age, max_age))
incluir_sin_edad = st.sidebar.checkbox("Incluir registros sin edad", value=True)

# Controles de gráficos
top_paises_n = st.sidebar.slider("Top países (N)", min_value=5, max_value=30, value=10)
bins_edad = st.sidebar.slider("Histograma edades: bins", min_value=5, max_value=30, value=10)
hist_show_no_age = st.sidebar.checkbox("Histograma: mostrar barra 'Sin edad'", value=False)


# Aplicar filtros
mask = pd.Series([True] * len(df))
if query:
    mask &= df["Nombre completo"].fillna("").str.contains(query.strip(), case=False, na=False)
if id_query and id_col:
    mask &= df[id_col].astype(str).fillna("").str.contains(id_query.strip(), case=False, na=False)
if sexo_sel:
    mask &= df["Sexo"].astype(str).isin(sexo_sel)
if nivel_sel:
    mask &= df["Nivel académico"].astype(str).isin(nivel_sel)
if pais_sel and pais_col:
    mask &= df[pais_col].astype(str).str.strip().isin(pais_sel)

# Filtro por idiomas y nivel (match ANY)
if lang_cols and (lang_sel or level_sel):
    # columnas candidatas en función de selección
    selected_cols = [
        c for (c, L, V) in lang_cols
        if (not lang_sel or L in lang_sel) and (not level_sel or V in level_sel)
    ]
    if selected_cols:
        any_lang = pd.Series([False] * len(df))
        for c in selected_cols:
            s = df[c].astype(str).str.strip()
            any_lang |= (~s.eq("") & ~df[c].isna())
        mask &= any_lang
areas_mask = filter_by_areas(df, areas_sel, match_all)
mask &= areas_mask

if edad_col in df.columns:
    age_series = pd.to_numeric(df[edad_col], errors="coerce")
    valid_age_mask = (age_series >= age_range[0]) & (age_series <= age_range[1])
    if incluir_sin_edad:
        mask &= (valid_age_mask | age_series.isna())
    else:
        mask &= valid_age_mask

df_filtered = df[mask].copy()


# KPIs
# KPIs
def kpi_card(title, value):
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{title}</div>
        <div class="metric-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4, gap="medium")
with col1:
    kpi_card("Voluntarios", f"{len(df_filtered):,}")
with col2:
    con_areas = df_filtered.get("Áreas de interés (count)", pd.Series([None]*len(df_filtered))).fillna(0)
    kpi_card("Con Áreas", f"{int((con_areas > 0).sum()):,}")
with col3:
    kpi_card("Sexos Distintos", df_filtered.get("Sexo", pd.Series()).nunique())
with col4:
    kpi_card("Niveles Acad.", df_filtered.get("Nivel académico", pd.Series()).nunique())


# Distribuciones
# Helper para gráficos consistentes
def make_simple_bar(df_in, x_col, y_col, title, sort_x=None):
    c = alt.Chart(df_in).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
        x=alt.X(x_col, sort=sort_x, axis=alt.Axis(labelAngle=-45, title=None)),
        y=alt.Y(y_col, axis=alt.Axis(title=None, grid=False)),
        tooltip=[x_col, y_col],
        color=alt.value("#0F9D58")
    ).properties(
        title=title,
        height=250
    ).configure_view(
        strokeWidth=0
    )
    return c

# Distribuciones
st.subheader("Distribuciones")
cols = st.columns(3, gap="medium")
with cols[0]:
    if "Sexo" in df_filtered.columns:
        counts = df_filtered["Sexo"].value_counts().reset_index()
        counts.columns = ["Sexo", "Cantidad"]
        st.altair_chart(
            make_simple_bar(counts, "Sexo", "Cantidad", "Por Sexo"), 
            use_container_width=True
        )
with cols[1]:
    if "Nivel académico" in df_filtered.columns:
        counts = df_filtered["Nivel académico"].value_counts().reset_index()
        counts.columns = ["Nivel", "Cantidad"]
        st.altair_chart(
            make_simple_bar(counts, "Nivel", "Cantidad", "Por Nivel Académico"),
            use_container_width=True
        )
with cols[2]:
    if areas_col in df_filtered.columns:
        top_areas = (
            df_filtered[areas_col]
            .dropna()
            .astype(str)
            .str.split(";")
            .explode()
            .str.strip()
            .value_counts()
            .head(10)
            .reset_index()
        )
        top_areas.columns = ["Área", "Cantidad"]
        st.altair_chart(
            make_simple_bar(top_areas, "Área", "Cantidad", "Top 10 Áreas de Interés", sort_x="-y"),
            use_container_width=True
        )

# Gráficos adicionales
# Gráficos adicionales
st.subheader("Gráficos adicionales")
colA, colB = st.columns(2, gap="medium")
with colA:
    # Top países (normalizado si existe)
    if pais_col and pais_col in df_filtered.columns:
        pais_series = (
            df_filtered[pais_col]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        pais_series = pais_series[pais_series != ""]
        top_paises = pais_series.value_counts().head(top_paises_n).reset_index()
        top_paises.columns = ["País", "Cantidad"]
        
        if len(top_paises) > 0:
            st.altair_chart(
                make_simple_bar(top_paises, "País", "Cantidad", f"Top {top_paises_n} Países", sort_x="-y"),
                use_container_width=True
            )
        else:
            st.info("No hay países para mostrar.")
    else:
        st.warning("Columna de país no disponible.")

with colB:
    # Histograma de edades con control de bins y toggle de 'sin edad'
    age_series = pd.to_numeric(df_filtered.get(edad_col, pd.Series()), errors="coerce")
    age_valid = age_series.dropna()
    
    if len(age_valid) > 0:
        chart_data = pd.DataFrame({"Edad": age_valid})
        
        # Histograma Altair
        base = alt.Chart(chart_data)
        hist = base.mark_bar(color="#0F9D58").encode(
            x=alt.X("Edad", bin=alt.Bin(maxbins=bins_edad), title="Rango de Edad"),
            y=alt.Y('count()', title="Voluntarios"),
            tooltip=[alt.Tooltip("Edad", bin=True), 'count()']
        ).properties(title="Distribución de Edad")
        
        st.altair_chart(hist, use_container_width=True)
    else:
        st.info("No hay datos de edad válidos para el histograma.")

# Idiomas: gráfico por nivel
st.subheader("Idiomas")
if lang_cols:
    # contar por idioma y nivel en df_filtrado
    df_lang = df_filtered.copy()
    counts = {}
    for c, L, V in lang_cols:
        s = df_lang[c].astype(str).str.strip()
        cnt = int((~s.eq("") & ~df_lang[c].isna()).sum())
        if cnt > 0:
            counts[(L, V)] = counts.get((L, V), 0) + cnt
            
    if counts:
        # Preparar DataFrame largo para Altair
        data_list = [{"Idioma": L, "Nivel": V, "Cantidad": C} for (L, V), C in counts.items()]
        chart_df = pd.DataFrame(data_list)
        
        # Orden de niveles
        level_order = ["Básico", "Intermedio", "Avanzado"]
        
        c = alt.Chart(chart_df).mark_bar().encode(
            x=alt.X("Idioma", sort="-y"),
            y=alt.Y("Cantidad"),
            color=alt.Color("Nivel", scale=alt.Scale(domain=level_order, scheme="greens")),
            tooltip=["Idioma", "Nivel", "Cantidad"]
        ).properties(
            title="Idiomas por Nivel",
            height=300
        )
        
        st.altair_chart(c, use_container_width=True)
    else:
        st.info("No hay datos de idiomas para mostrar.")
else:
    st.write("Columnas de idiomas no detectadas.")


# Tabla
st.subheader("Datos filtrados")
df_display = ensure_arrow_compatible(df_filtered)
st.dataframe(df_display, width="stretch")


# Exportación
st.subheader("Exportar")
csv_bytes = df_filtered.to_csv(index=False).encode("utf-8")
st.download_button(
    label="Descargar CSV",
    data=csv_bytes,
    file_name="voluntariado_filtrado.csv",
    mime="text/csv",
)

excel_bytes = to_excel_bytes(df_filtered, sheet_name="Merged")
st.download_button(
    label="Descargar Excel",
    data=excel_bytes,
    file_name="voluntariado_filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)