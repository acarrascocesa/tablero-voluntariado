import pandas as pd

PATH = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
SHEET = "Merged"

COLS = [
    "País (normalizado)",
    "País de residencia",
    "País",
    "Pais",
    "Country",
]


def is_blank_series(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    return s.eq("") | series.isna()


def main():
    df = pd.read_excel(PATH, sheet_name=SHEET)
    found = [c for c in COLS if c in df.columns]
    print("Columnas país presentes:", found)
    for c in found:
        blanks = is_blank_series(df[c])
        print(
            f"Columna: {c} | Total filas: {len(df)} | En blanco: {int(blanks.sum())} | No en blanco: {int((~blanks).sum())}"
        )


if __name__ == "__main__":
    main()