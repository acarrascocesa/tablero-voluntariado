import re
import unicodedata
import pandas as pd


DEDUPE_PATH = "Voluntariado Base + WPForms - Areas (Dedup Nombre).xlsx"
SHEET = "Merged"

PAIS_CANDIDATES = [
    "País",
    "Pais",
    "País de residencia",
    "País (Residencia)",
    "Country",
]


def normalize_country(value: str) -> str:
    s = str(value).strip()
    # Collapse internal whitespace
    s = re.sub(r"\s+", " ", s)
    # Lowercase
    s = s.lower()
    # Remove accents
    s = (
        unicodedata.normalize("NFKD", s)
        .encode("ascii", "ignore")
        .decode("ascii")
    )
    # Remove punctuation
    s = re.sub(r"[^a-z0-9 ]", "", s)
    # Trim again
    s = s.strip()
    return s


def main():
    try:
        df = pd.read_excel(DEDUPE_PATH, sheet_name=SHEET)
    except Exception as e:
        print(f"Error leyendo {DEDUPE_PATH}: {e}")
        return

    pais_col = next((c for c in PAIS_CANDIDATES if c in df.columns), None)
    if not pais_col:
        print("No se encontró una columna de país entre candidatos:", ", ".join(PAIS_CANDIDATES))
        return

    raw = df[pais_col].fillna("").astype(str).str.strip()
    unique_raw = sorted(set([x for x in raw.unique() if x]))
    print(f"Columna de país detectada: {pais_col}")
    print(f"Valores únicos (raw): {len(unique_raw)}")

    # Normalizar y agrupar variantes
    norm = raw.apply(normalize_country)
    df_out = pd.DataFrame({"pais_raw": raw, "pais_norm": norm})
    grouped = df_out.groupby("pais_norm")

    unique_norm = grouped.ngroups - (1 if "" in df_out["pais_norm"].unique() else 0)
    print(f"Valores únicos normalizados (sin acentos/puntuación/caso): {unique_norm}")

    # Top 20 grupos por tamaño, mostrando algunas variantes crudas
    sizes = grouped.size().sort_values(ascending=False)
    print("\nTop 20 países (normalizados) por frecuencia y ejemplos de variantes:")
    for norm_value, count in sizes.head(20).items():
        if norm_value == "":
            continue
        variants = (
            grouped.get_group(norm_value)["pais_raw"]
            .value_counts()
            .index.tolist()
        )
        examples = ", ".join(variants[:4])
        print(f"- '{norm_value}' -> {count} filas; ejemplos: {examples}")

    # Exportar CSV para revisión completa
    df_out.to_csv("pais_variantes.csv", index=False)
    print("\nExportado detalle a pais_variantes.csv (columns: pais_raw, pais_norm)")


if __name__ == "__main__":
    main()