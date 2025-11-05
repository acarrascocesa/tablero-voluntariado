import re
import pandas as pd

SHEET = "Merged"
PATHS = [
    "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx",
    "Voluntariado Base + WPForms - Areas (Dedup Nombre).xlsx",
]

PATTERN = re.compile(r"id(i|í)oma|idiomas|language|languages", re.IGNORECASE)


def main():
    for path in PATHS:
        try:
            df = pd.read_excel(path, sheet_name=SHEET)
        except Exception as e:
            print("Error leyendo", path, e)
            continue
        cols = [str(c) for c in df.columns]
        cand = [c for c in cols if PATTERN.search(c)]
        print("\nArchivo:", path)
        print("Columnas encontradas:", cand)
        for c in cand[:5]:
            vals = df[c].dropna().astype(str).unique()[:5]
            print(f"- {c}: ejemplos -> {list(vals)}")


if __name__ == "__main__":
    main()