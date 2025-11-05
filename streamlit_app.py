import datetime
import numpy as np
from io import BytesIO
from typing import Optional, Tuple, List

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Tablero Voluntariado", layout="wide")
st.title("Tablero de Voluntariado")


@st.cache_data(show_spinner=False)
def load_data() -> Tuple[pd.DataFrame, Optional[str]]:
    path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
    try:
        df = pd.read_excel(path, sheet_name="Merged")
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
df, source_path = load_data()
if source_path is None or df.empty:
    st.error("No se encontró el archivo deduplicado por nombre: 'Voluntariado Base + WPForms - Areas (Dedup Nombre).xlsx'.")
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
st.sidebar.header("Controles")

# Búsqueda por nombre
query = st.sidebar.text_input("Buscar por nombre")

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
if sexo_sel:
    mask &= df["Sexo"].astype(str).isin(sexo_sel)
if nivel_sel:
    mask &= df["Nivel académico"].astype(str).isin(nivel_sel)
if pais_sel and pais_col:
    mask &= df[pais_col].astype(str).str.strip().isin(pais_sel)
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
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Voluntarios (filtrados)", len(df_filtered))
with col2:
    con_areas = df_filtered.get("Áreas de interés (count)", pd.Series([None]*len(df_filtered))).fillna(0)
    st.metric("Con áreas marcadas", int((con_areas > 0).sum()))
with col3:
    st.metric("Sexo distintos", df_filtered.get("Sexo", pd.Series()).nunique())
with col4:
    st.metric("Nivel académico distintos", df_filtered.get("Nivel académico", pd.Series()).nunique())


# Distribuciones
st.subheader("Distribuciones")
cols = st.columns(3)
with cols[0]:
    if "Sexo" in df_filtered.columns:
        st.bar_chart(df_filtered["Sexo"].value_counts().sort_index())
with cols[1]:
    if "Nivel académico" in df_filtered.columns:
        st.bar_chart(df_filtered["Nivel académico"].value_counts().sort_index())
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
        )
        st.bar_chart(top_areas)

# Gráficos adicionales
st.subheader("Gráficos adicionales")
colA, colB = st.columns(2)
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
        top_paises = pais_series.value_counts().head(top_paises_n)
        if len(top_paises) > 0:
            st.bar_chart(top_paises)
        else:
            st.write("No hay países para mostrar.")
    else:
        st.write("Columna de país no disponible.")

with colB:
    # Histograma de edades con control de bins y toggle de 'sin edad'
    age_series = pd.to_numeric(df_filtered.get(edad_col, pd.Series()), errors="coerce")
    age_valid = age_series.dropna()
    if len(age_valid) > 0:
        # Construir histograma
        min_age_val = int(np.floor(age_valid.min()))
        max_age_val = int(np.ceil(age_valid.max()))
        counts, bin_edges = np.histogram(age_valid, bins=bins_edad, range=(min_age_val, max_age_val))
        labels = [f"{int(bin_edges[i])}–{int(bin_edges[i+1])}" for i in range(len(bin_edges)-1)]
        hist_series = pd.Series(counts, index=labels)
        # Agregar barra 'Sin edad' si se solicita
        if hist_show_no_age:
            sin_edad_count = int(age_series.isna().sum())
            hist_series = pd.concat([hist_series, pd.Series({"Sin edad": sin_edad_count})])
        st.bar_chart(hist_series)
    else:
        if hist_show_no_age:
            sin_edad_count = int(age_series.isna().sum())
            st.bar_chart(pd.Series({"Sin edad": sin_edad_count}))
        else:
            st.write("No hay datos de edad válidos para el histograma.")


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