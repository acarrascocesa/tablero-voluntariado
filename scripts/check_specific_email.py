
import pandas as pd
import sys

# Config
master_path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
target_email = "batistanero7@gmail.com"

def norm_email(x):
    s = str(x).strip().lower()
    return s if s and s != "nan" else None

try:
    print(f"Leyendo master: {master_path}")
    df = pd.read_excel(master_path)
    
    # Identify email column
    email_cols = [c for c in df.columns if "correo" in c.lower() or "email" in c.lower()]
    
    if not email_cols:
        print("No se encontraron columnas de email.")
        sys.exit(1)
        
    print(f"Columnas de email encontradas: {email_cols}")
    
    total_matches = 0
    matches_details = []

    for col in email_cols:
        # Normalize and check
        matches = df[df[col].apply(norm_email) == target_email]
        count = len(matches)
        if count > 0:
            total_matches += count
            matches_details.append((col, matches.index.tolist()))
            
            # Print details of matches
            print(f"\nEncontrados {count} en columna '{col}':")
            for idx, row in matches.iterrows():
                # Print some identifying info like Name if available
                name_cols = [c for c in df.columns if "nombre" in c.lower()]
                name_val = row[name_cols[0]] if name_cols else "N/A"
                print(f"  - Fila {idx+2} (Excel): {name_val}") # +2 assuming header is row 1 and 0-index

    print(f"\nTotal de coincidencias para '{target_email}': {total_matches}")

except Exception as e:
    print(f"Error: {e}")
