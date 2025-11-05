import re
import sys
from typing import List

import pandas as pd


def is_area_column(col_name: str) -> bool:
    if col_name is None:
        return False
    s = str(col_name)
    return "Área(s) de Interés para Voluntariado" in s


def extract_area_label(col_name: str) -> str:
    s = str(col_name)
    # Expected format: "Área(s) de Interés para Voluntariado: <Label>"
    parts = s.split(":", 1)
    label = parts[1].strip() if len(parts) > 1 else s
    # Normalize whitespace and quotes/newlines
    label = re.sub(r"\s+", " ", label)
    label = label.strip().strip('"')
    return label


def main(input_xlsx: str, output_xlsx: str) -> None:
    df = pd.read_excel(input_xlsx, sheet_name="Merged")

    area_cols: List[str] = [c for c in df.columns if is_area_column(c)]
    area_labels = {c: extract_area_label(c) for c in area_cols}

    def collect_areas(row) -> List[str]:
        selected = []
        for c in area_cols:
            val = row.get(c)
            # WPForms marks checkbox as "Checked"; also accept truthy non-empty
            if isinstance(val, str):
                v = val.strip().lower()
                if v == "checked" or v == "true" or v == "sí" or v == "si":
                    selected.append(area_labels[c])
                elif v != "" and v not in ("nan", "none"):
                    # Any non-empty string counts as selected
                    selected.append(area_labels[c])
            elif pd.notna(val) and val:
                selected.append(area_labels[c])
        # Deduplicate while preserving order
        seen = set()
        uniq = []
        for a in selected:
            if a not in seen:
                seen.add(a)
                uniq.append(a)
        return uniq

    areas_list = df.apply(collect_areas, axis=1)
    df["Áreas de interés (lista)"] = areas_list.apply(lambda xs: "; ".join(xs))
    df["Áreas de interés (count)"] = areas_list.apply(len)

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged")

    # Summary
    total = len(df)
    with_areas = int((df["Áreas de interés (count)"] > 0).sum())
    print("Input:", input_xlsx)
    print("Output:", output_xlsx)
    print("Filas:", total, "| Con al menos un área:", with_areas)
    print("Columnas de áreas detectadas:", len(area_cols))
    print("Ejemplos de etiquetas:", list(set(area_labels.values()))[:10])


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: .venv/bin/python normalize_areas_interes.py 'Voluntariado Base + WPForms.xlsx' 'Voluntariado Base + WPForms - Areas.xlsx'")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])