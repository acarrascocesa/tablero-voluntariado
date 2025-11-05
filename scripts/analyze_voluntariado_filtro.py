import re
import sys
from typing import List

import pandas as pd


def main(path: str) -> None:
    xls = pd.ExcelFile(path)
    print("Sheets:", xls.sheet_names)
    df = pd.read_excel(xls, xls.sheet_names[0])
    cols: List[str] = [str(c) for c in df.columns]
    print("Filas x Columnas:", df.shape)
    print("Columnas (primeras 40):")
    for c in cols[:40]:
        print("-", c)

    # Detect candidate keys for joining
    key_patterns = r"correo|email|ident|cedul|tel[eé]fono|name|nombre"
    keys = [c for c in cols if re.search(key_patterns, c, re.IGNORECASE)]
    print("\nPosibles llaves de cruce:")
    for k in keys:
        print("-", k)

    print("\nMuestras de valores (hasta 5 por columna):")
    for k in keys[:10]:
        vals = df[k].dropna().astype(str).unique()[:5]
        print(f"{k} -> {list(vals)}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: .venv/bin/python analyze_voluntariado_filtro.py 'Voluntariado Filtro - All Data.xlsx'")
        sys.exit(1)
    main(sys.argv[1])