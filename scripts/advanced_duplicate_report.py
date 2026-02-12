import pandas as pd
import re
from datetime import datetime

def normalize_email(x):
    s = str(x).strip().lower()
    return s if s and s != "nan" and "@" in s else None

def normalize_phone(x):
    digits = re.sub(r"\D+", "", str(x))
    # Standardize to last 10 digits if longer (common for RD/International)
    if digits and len(digits) >= 7:
        return digits[-10:] if len(digits) > 10 else digits
    return None

def normalize_id(x):
    s = re.sub(r"\W+", "", str(x)).lower()
    # Avoid generic values
    if s in ["cedula", "pasaporte", "nan", "none", "n/a", "no", "0", "porasignar"]:
        return None
    return s if s and len(s) > 4 else None

def normalize_str(s):
    if pd.isna(s): return ""
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def normalize_dob(x):
    try:
        if pd.isna(x): return None
        ts = pd.to_datetime(x, dayfirst=True, errors="coerce")
        return ts.date() if pd.notna(ts) else None
    except Exception:
        return None

def create_report(master_path, output_path):
    print(f"Leyendo maestro: {master_path}...")
    df = pd.read_excel(master_path)
    df['original_row_index'] = df.index + 2  # +1 for 0-index, +1 for header
    
    cols = df.columns
    email_col = next((c for c in cols if "correo" in c.lower() or "email" in c.lower()), None)
    phone_col = next((c for c in cols if "tel" in c.lower() or "phone" in c.lower()), None)
    id_col = next((c for c in cols if "identi" in c.lower() or "cedula" in c.lower()), None)
    
    name_cols = [c for c in cols if "nombre" in c.lower()]
    full_name_col = next((c for c in name_cols if "completo" in c.lower() and "first" not in c.lower() and "last" not in c.lower()), None)
    first_col = next((c for c in name_cols if "first" in c.lower()), None)
    last_col = next((c for c in name_cols if "last" in c.lower()), None)
    dob_col = next((c for c in cols if "nacimiento" in c.lower() or "birth" in c.lower()), None)

    # Prepare keys
    print("Normalizando llaves para análisis...")
    if email_col: df['_key_email'] = df[email_col].apply(normalize_email)
    if phone_col: df['_key_phone'] = df[phone_col].apply(normalize_phone)
    if id_col: df['_key_id'] = df[id_col].apply(normalize_id)
    
    def get_full_name(row):
        if full_name_col and pd.notna(row[full_name_col]):
            return normalize_str(row[full_name_col])
        elif first_col and last_col:
            return normalize_str(f"{row[first_col]} {row[last_col]}")
        return ""

    df['_key_name'] = df.apply(get_full_name, axis=1)
    if dob_col:
        df['_key_dob'] = df[dob_col].apply(normalize_dob)
        df['_key_name_dob'] = df.apply(lambda r: f"{r['_key_name']}|{r['_key_dob']}" if r['_key_name'] and r['_key_dob'] else None, axis=1)
    else:
        df['_key_name_dob'] = None

    # Find duplicates
    def get_dups(df, key_col):
        if key_col not in df.columns: return pd.DataFrame()
        valid_keys = df[df[key_col].notna()]
        duplicates = valid_keys[valid_keys.duplicated(subset=[key_col], keep=False)]
        return duplicates.sort_values(by=[key_col, 'original_row_index'])

    print("Identificando grupos de duplicados...")
    dup_email = get_dups(df, '_key_email')
    dup_phone = get_dups(df, '_key_phone')
    dup_id = get_dups(df, '_key_id')
    dup_name_dob = get_dups(df, '_key_name_dob')

    # New: Combined Name + ID
    df['_key_name_id'] = df.apply(lambda r: f"{r['_key_name']}|{r['_key_id']}" if r['_key_name'] and r['_key_id'] else None, axis=1)
    dup_name_id = get_dups(df, '_key_name_id')

    # Summary
    summary_data = {
        "Categoría": ["Email", "Teléfono", "ID (Cédula/Pasaporte)", "Nombre + Fecha Nacimiento", "Nombre + ID"],
        "Grupos de Duplicados": [
            dup_email['_key_email'].nunique() if not dup_email.empty else 0,
            dup_phone['_key_phone'].nunique() if not dup_phone.empty else 0,
            dup_id['_key_id'].nunique() if not dup_id.empty else 0,
            dup_name_dob['_key_name_dob'].nunique() if not dup_name_dob.empty else 0,
            dup_name_id['_key_name_id'].nunique() if not dup_name_id.empty else 0,
        ],
        "Total Filas Afectadas": [
            len(dup_email),
            len(dup_phone),
            len(dup_id),
            len(dup_name_dob),
            len(dup_name_id)
        ]
    }
    df_summary = pd.DataFrame(summary_data)

    # Save to Excel
    print(f"Guardando reporte en: {output_path}...")
    print("\n--- RESUMEN DE DUPLICADOS ENCONTRADOS ---")
    print(df_summary.to_string(index=False))
    print("------------------------------------------\n")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_summary.to_sheet_name = "Resumen"
        df_summary.to_excel(writer, sheet_name="Resumen", index=False)
        
        # Helper to clean temp keys before saving
        def clean_for_save(df_dup):
            cols_to_drop = [c for c in df_dup.columns if c.startswith('_key_')]
            return df_dup.drop(columns=cols_to_drop)

        if not dup_email.empty: clean_for_save(dup_email).to_excel(writer, sheet_name="Duplicados_Email", index=False)
        if not dup_phone.empty: clean_for_save(dup_phone).to_excel(writer, sheet_name="Duplicados_Telefono", index=False)
        if not dup_id.empty: clean_for_save(dup_id).to_excel(writer, sheet_name="Duplicados_ID", index=False)
        if not dup_name_dob.empty: clean_for_save(dup_name_dob).to_excel(writer, sheet_name="Duplicados_Nombre_Fecha", index=False)
        if not dup_name_id.empty: clean_for_save(dup_name_id).to_excel(writer, sheet_name="Duplicados_Nombre_ID", index=False)

    print("Reporte generado con éxito.")

if __name__ == "__main__":
    MASTER = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
    OUTPUT = "Reporte_Detallado_Duplicados.xlsx"
    create_report(MASTER, OUTPUT)
