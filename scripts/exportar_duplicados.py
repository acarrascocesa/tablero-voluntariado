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
            s = re.sub(r"[^0-9+]", "", s)
            if s.count('+') > 1:
                s = s.replace('+', '')
            return s
        return df[pcol].fillna('').astype(str).map(clean)
    return pd.Series([''] * len(df))


def normalize_id(df: pd.DataFrame) -> pd.Series:
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
        s = str(x).strip()
        s = re.sub(r"\s+", "", s)
        s = s.replace('"', '').lower()
        s = re.sub(r"[^0-9a-z-]", "", s)
        return s
    return series.fillna('').map(clean)


def build_duplicates_sheet(df: pd.DataFrame, key_series: pd.Series, key_label: str) -> pd.DataFrame:
    counts = Counter(key_series)
    dup_keys = {k: c for k, c in counts.items() if k and c > 1}
    mask = key_series.map(lambda k: k in dup_keys)
    out = df.loc[mask].copy()
    out.insert(0, 'Duplicate Key Type', key_label)
    out.insert(1, 'Duplicate Key Value', key_series.loc[mask].values)
    out.insert(2, 'Duplicate Group Count', key_series.loc[mask].map(lambda k: dup_keys.get(k, 0)).values)
    return out


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    keys = [
        ('Nombre', normalize_name(df)),
        ('Email', normalize_email(df)),
        ('Teléfono', normalize_phone(df)),
        ('Identificación', normalize_id(df)),
    ]
    for label, series in keys:
        counts = Counter(series)
        dup_groups = {k: c for k, c in counts.items() if k and c > 1}
        rows.append({
            'Clave': label,
            'Grupos duplicados': len(dup_groups),
            'Filas en grupos duplicados': sum(dup_groups.values()),
        })
    return pd.DataFrame(rows)


def main():
    input_path = 'Voluntariado Base + WPForms.xlsx'
    df = pd.read_excel(input_path, sheet_name='Merged')

    name_sheet = build_duplicates_sheet(df, normalize_name(df), 'Nombre completo')
    email_sheet = build_duplicates_sheet(df, normalize_email(df), 'Correo electrónico')
    phone_sheet = build_duplicates_sheet(df, normalize_phone(df), 'Teléfono')
    id_sheet = build_duplicates_sheet(df, normalize_id(df), 'Identificación')
    summary = build_summary(df)

    output_path = 'Voluntariado Duplicados.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        summary.to_excel(writer, index=False, sheet_name='Resumen')
        name_sheet.to_excel(writer, index=False, sheet_name='Duplicados Nombre')
        email_sheet.to_excel(writer, index=False, sheet_name='Duplicados Email')
        phone_sheet.to_excel(writer, index=False, sheet_name='Duplicados Teléfono')
        id_sheet.to_excel(writer, index=False, sheet_name='Duplicados ID')

    print('Archivo de duplicados generado:', output_path)
    print('Hojas:', ['Resumen', 'Duplicados Nombre', 'Duplicados Email', 'Duplicados Teléfono', 'Duplicados ID'])
    print('Filas por hoja:', {
        'Resumen': len(summary),
        'Duplicados Nombre': len(name_sheet),
        'Duplicados Email': len(email_sheet),
        'Duplicados Teléfono': len(phone_sheet),
        'Duplicados ID': len(id_sheet),
    })


if __name__ == '__main__':
    main()