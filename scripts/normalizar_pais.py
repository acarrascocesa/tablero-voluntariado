import re
import unicodedata
import pandas as pd


INPUT_XLSX = "Voluntariado Base + WPForms - Areas (Dedup Nombre).xlsx"
OUTPUT_XLSX = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
SHEET = "Merged"

PAIS_CANDIDATES = [
    "País (normalizado)",  # por si ya existe
    "País de residencia",
    "País",
    "Pais",
    "País (Residencia)",
    "Country",
]


def strip_accents_lower(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = (
        unicodedata.normalize("NFKD", s)
        .encode("ascii", "ignore")
        .decode("ascii")
    )
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return s.strip()


# Equivalencias y canonización
CANON_EQUIV = {
    # Unificar sinónimos/abreviaturas
    "usa": "Estados Unidos",
    "eeuu": "Estados Unidos",
    "estados unidos": "Estados Unidos",
    # Países con acento/canon específico
    "republica dominicana": "República Dominicana",
    "mexico": "México",
    "peru": "Perú",
    "espana": "España",
    "canada": "Canadá",
}


def title_case_ascii(norm: str) -> str:
    # Título simple para valores no mapeados (sin acentos)
    return " ".join(w.capitalize() for w in norm.split())


def canonical_country(raw: str) -> str:
    norm = strip_accents_lower(raw)
    if not norm:
        return ""
    if norm in CANON_EQUIV:
        return CANON_EQUIV[norm]
    # Valores comunes que no requieren acentos extra
    return title_case_ascii(norm)


def main():
    try:
        df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET)
    except Exception as e:
        print(f"Error leyendo {INPUT_XLSX}: {e}")
        return

    pais_col = next((c for c in PAIS_CANDIDATES if c in df.columns), None)
    if not pais_col:
        print("No se encontró columna de país. Candidatos:", ", ".join(PAIS_CANDIDATES))
        return

    raw = df[pais_col].fillna("").astype(str).str.strip()
    df["País (normalizado)"] = raw.apply(canonical_country)

    # Guardar a nuevo Excel
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=SHEET)

    unique_raw = len(set([x for x in raw.unique() if x]))
    unique_norm = len(set([x for x in df["País (normalizado)"].unique() if x]))
    print("Input:", INPUT_XLSX)
    print("Output:", OUTPUT_XLSX)
    print("Columna usada:", pais_col)
    print("Filas:", len(df))
    print("Valores país (raw) únicos:", unique_raw)
    print("Valores país (normalizado) únicos:", unique_norm)


if __name__ == "__main__":
    main()