
import pandas as pd
from collections import Counter
import re

master_path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"

def norm_email(x):
    s = str(x).strip().lower()
    return s if s and s != "nan" else None

def norm_phone(x):
    digits = re.sub(r"\D+", "", str(x))
    return digits if digits and len(digits) >= 7 else None

def norm_id(x):
    s = re.sub(r"\W+", "", str(x)).lower()
    return s if s and s != "nan" else None

def norm_str(s):
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def norm_dob(x):
    try:
        ts = pd.to_datetime(x, dayfirst=True, errors="coerce")
        return ts.date() if pd.notna(ts) else None
    except Exception:
        return None

def analyze(df):
    print(f"Total filas: {len(df)}")
    
    # Identify columns
    cols = df.columns
    email_col = next((c for c in cols if "correo" in c.lower() or "email" in c.lower()), None)
    phone_col = next((c for c in cols if "tel" in c.lower() or "phone" in c.lower()), None)
    id_col = next((c for c in cols if "identi" in c.lower() or "cedula" in c.lower()), None)
    
    name_cols = [c for c in cols if "nombre" in c.lower()]
    full_name_col = next((c for c in name_cols if "completo" in c.lower() and "first" not in c.lower() and "last" not in c.lower()), None)
    first_col = next((c for c in name_cols if "first" in c.lower()), None)
    last_col = next((c for c in name_cols if "last" in c.lower()), None)
    
    dob_col = next((c for c in cols if "nacimiento" in c.lower() or "birth" in c.lower()), None)

    # 1. Check Emails
    if email_col:
        emails = df[email_col].apply(norm_email).dropna()
        dup_emails = [k for k, v in Counter(emails).items() if v > 1]
        print(f"\n[EMAIL] Duplicados encontrados: {len(dup_emails)}")
        if dup_emails:
            print(f"Ejemplos: {dup_emails[:5]}")
            
    # 2. Check Phones
    if phone_col:
        phones = df[phone_col].apply(norm_phone).dropna()
        dup_phones = [k for k, v in Counter(phones).items() if v > 1]
        print(f"\n[TELÉFONO] Duplicados encontrados: {len(dup_phones)}")
        if dup_phones:
            print(f"Ejemplos: {dup_phones[:5]}")

    # 3. Check IDs
    if id_col:
        ids = df[id_col].apply(norm_id).dropna()
        dup_ids = [k for k, v in Counter(ids).items() if v > 1]
        print(f"\n[ID] Duplicados encontrados: {len(dup_ids)}")
        if dup_ids:
            print(f"Ejemplos: {dup_ids[:5]}")

    # 4. Check Name + DOB
    names_dob = []
    for _, row in df.iterrows():
        nm = None
        if full_name_col and pd.notna(row[full_name_col]):
            nm = norm_str(row[full_name_col])
        elif first_col and last_col:
             nm = norm_str(f"{row[first_col]} {row[last_col]}")
        
        dob = norm_dob(row[dob_col]) if dob_col else None
        
        if nm and dob:
            names_dob.append((nm, dob))
            
    dup_namedob = [k for k, v in Counter(names_dob).items() if v > 1]
    print(f"\n[NOMBRE + FECHA NAC] Duplicados encontrados: {len(dup_namedob)}")
    if dup_namedob:
        print(f"Ejemplos: {dup_namedob[:5]}")

if __name__ == "__main__":
    try:
        df = pd.read_excel(master_path)
        analyze(df)
    except Exception as e:
        print(f"Error: {e}")
