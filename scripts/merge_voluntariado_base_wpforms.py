import sys
import pandas as pd


def make_fullname_key(df: pd.DataFrame, first_col: str, last_col: str) -> pd.Series:
    f = df[first_col].fillna("").astype(str).str.strip()
    l = df[last_col].fillna("").astype(str).str.strip()
    return (f + " " + l).str.lower()


def main(base_path: str, wpforms_csv_path: str, output_xlsx_path: str) -> None:
    # Read base (Excel)
    xls = pd.ExcelFile(base_path)
    sheet_name = xls.sheet_names[0]
    df_base = pd.read_excel(xls, sheet_name)

    # Read WPForms (CSV)
    df_wp = pd.read_csv(wpforms_csv_path)

    # Build keys
    base_key = make_fullname_key(df_base, "Nombre completo: First", "Nombre completo: Last")
    wp_key = make_fullname_key(df_wp, "Nombre completo: First", "Nombre completo: Last")

    df_base["__key__"] = base_key
    df_wp["__key__"] = wp_key

    # Determine columns to carry over from WPForms that are missing in base
    base_cols = set(df_base.columns.tolist())
    wp_extra_cols = [c for c in df_wp.columns.tolist() if c not in base_cols and c != "__key__"]

    df_wp_extra = df_wp[["__key__"] + wp_extra_cols]

    # Left merge: keep base as authoritative, add extra columns from WPForms where matched
    df_merged = df_base.merge(df_wp_extra, how="left", on="__key__")

    # Drop technical key
    df_merged = df_merged.drop(columns=["__key__"]) if "__key__" in df_merged.columns else df_merged

    # Write Excel output
    with pd.ExcelWriter(output_xlsx_path, engine="openpyxl") as writer:
        df_merged.to_excel(writer, index=False, sheet_name="Merged")

    # Print summary
    print("Base sheet:", sheet_name)
    print("Base rows:", len(df_base))
    print("WPForms rows:", len(df_wp))
    print("Merged rows:", len(df_merged))
    print("Columns added from WPForms:", len(wp_extra_cols))
    print("Sample added columns:", wp_extra_cols[:10])
    print("Output:", output_xlsx_path)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Uso: .venv/bin/python merge_voluntariado_base_wpforms.py 'Voluntariado Filtro - All Data.xlsx' voluntarios.csv 'Voluntariado Base + WPForms.xlsx'")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])