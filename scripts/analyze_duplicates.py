from collections import Counter
import re
import pandas as pd


def normalize_name(df: pd.DataFrame) -> pd.Series:
    fcol = 'Nombre completo: First'
    lcol = 'Nombre completo: Last'
    if fcol in df.columns and lcol in df.columns:
        first = df[fcol].fillna('').astype(str).str.strip()
        last = df[lcol].fillna('').astype(str).str.strip()
        return (first + ' ' + last).str.lower()
    return pd.Series([''] * len(df))


def normalize_email(df: pd.DataFrame) -> pd.Series:
    ecol = 'Correo electrónico'
    if ecol in df.columns:
        return df[ecol].fillna('').astype(str).str.strip().str.lower()
    return pd.Series([''] * len(df))


def normalize_phone(df: pd.DataFrame) -> pd.Series:
    pcol = 'Teléfono'
    if pcol in df.columns:
        def clean(x: str) -> str:
            s = str(x)
            s = s.strip().lower().replace("'+", "+")
            # keep digits and plus
            s = re.sub(r"[^0-9+]", "", s)
            # if multiple plus, keep last occurrence only
            s = s.replace('+', '') if s.count('+') > 1 else s
            return s
        return df[pcol].fillna('').astype(str).map(clean)
    return pd.Series([''] * len(df))


def normalize_id(df: pd.DataFrame) -> pd.Series:
    # Try WPForms 'Identificación' first, fall back to base 'Unnamed: 8'
    wp_id = 'Identificación'
    base_id = 'Unnamed: 8'
    series = None
    if wp_id in df.columns:
        series = df[wp_id].astype(str)
    if series is None and base_id in df.columns:
        series = df[base_id].astype(str)
    if series is None:
        return pd.Series([''] * len(df))
    def clean(x: str) -> str:
        s = str(x)
        s = s.strip()
        s = re.sub(r"\s+", "", s)
        s = s.replace('"', '')
        s = s.lower()
        # keep digits, letters and dashes
        s = re.sub(r"[^0-9a-z-]", "", s)
        return s
    return series.fillna('').map(clean)


def summarize_duplicates(key_series: pd.Series, label: str) -> None:
    counts = Counter(key_series)
    dup_groups = {k: c for k, c in counts.items() if k and c > 1}
    total_groups = len(dup_groups)
    total_rows_dup = sum(c for c in dup_groups.values())
    print(f"\n[{label}] Grupos duplicados: {total_groups} | Filas en grupos duplicados: {total_rows_dup}")
    if total_groups:
        top = sorted(dup_groups.items(), key=lambda x: x[1], reverse=True)[:15]
        print("Top duplicados (valor, conteo):")
        for v, c in top:
            print(f"- {v} -> {c}")


def main():
    # Prefer the latest combined file; both have sheet 'Merged'
    path = 'Voluntariado Base + WPForms.xlsx'
    df = pd.read_excel(path, sheet_name='Merged')
    print(f"Archivo: {path} | Filas: {len(df)} | Columnas: {len(df.columns)}")

    name_key = normalize_name(df)
    email_key = normalize_email(df)
    phone_key = normalize_phone(df)
    id_key = normalize_id(df)

    summarize_duplicates(name_key, 'Nombre completo')
    summarize_duplicates(email_key, 'Correo electrónico')
    summarize_duplicates(phone_key, 'Teléfono')
    summarize_duplicates(id_key, 'Identificación')


if __name__ == '__main__':
    main()