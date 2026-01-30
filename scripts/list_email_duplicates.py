
import pandas as pd
from collections import Counter
import os

master_path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
output_path = "reporte_duplicados_email.xlsx"

def norm_email(x):
    s = str(x).strip().lower()
    return s if s and s != "nan" else None

def main():
    if not os.path.exists(master_path):
        print(f"No se encuentra el archivo: {master_path}")
        return

    print(f"Leyendo {master_path}...")
    df = pd.read_excel(master_path)
    
    # Identify columns
    cols = df.columns
    email_col = next((c for c in cols if "correo" in c.lower() or "email" in c.lower()), None)
    
    if not email_col:
        print("No se encontró columna de email.")
        return

    # Normalize emails
    df['__email_norm__'] = df[email_col].apply(norm_email)
    
    # Find duplicates
    counts = Counter(df['__email_norm__'].dropna())
    dup_emails = [k for k, v in counts.items() if v > 1]
    
    print(f"Total emails únicos duplicados: {len(dup_emails)}")
    
    if not dup_emails:
        print("No hay duplicados por email.")
        return

    # Filter rows with duplicate emails
    dup_rows = df[df['__email_norm__'].isin(dup_emails)].copy()
    
    # Sort by email to group them visually
    dup_rows = dup_rows.sort_values(by=['__email_norm__'])
    
    # Select useful columns for inspection
    name_cols = [c for c in cols if "nombre" in c.lower()]
    phone_col = next((c for c in cols if "tel" in c.lower() or "phone" in c.lower()), None)
    id_col = next((c for c in cols if "identi" in c.lower() or "cedula" in c.lower()), None)
    photo_col = next((c for c in cols if "foto" in c.lower()), None)
    
    # Prioritize 'Nombre completo' if exists
    display_cols = []
    if name_cols:
        display_cols.extend(name_cols[:2])
    if phone_col:
        display_cols.append(phone_col)
    if id_col:
        display_cols.append(id_col)
    if photo_col:
        display_cols.append(photo_col)
        
    print("\n--- Muestra de duplicados (Primeros 10 grupos) ---")
    current_email = None
    shown_groups = 0
    
    for idx, row in dup_rows.iterrows():
        email = row['__email_norm__']
        if email != current_email:
            if shown_groups >= 10:
                print("\n... (más grupos en el archivo exportado)")
                break
            current_email = email
            shown_groups += 1
            print(f"\n[Email: {email}]")
            
        # Print row info
        info = []
        for c in display_cols:
            val = str(row[c])
            # Truncate long values or URLs
            if "http" in val:
                val = "[TIENE FOTO]"
            elif len(val) > 30: 
                val = val[:27] + "..."
            info.append(f"{c}: {val}")
        print(f"  Fila {idx}: " + " | ".join(info))

    # Export to Excel (clean up temp col first)
    del dup_rows['__email_norm__']
    dup_rows.to_excel(output_path, index=True)
    print(f"\nReporte completo exportado a: {output_path}")

if __name__ == "__main__":
    main()
