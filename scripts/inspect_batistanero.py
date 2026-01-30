
import pandas as pd

master_path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
df = pd.read_excel(master_path)

target_email = "batistanero7@gmail.com"
email_cols = [c for c in df.columns if "correo" in c.lower() or "email" in c.lower()]

# Filter rows
matches = df[df[email_cols[0]].astype(str).str.strip().str.lower() == target_email].copy()

print("Columnas disponibles:", df.columns.tolist())
print("-" * 20)
print(f"Encontrados {len(matches)} registros.")

# Try to identify photo column
photo_keywords = ['foto', 'photo', 'imagen', 'image', 'picture']
photo_cols = [c for c in df.columns if any(k in c.lower() for k in photo_keywords)]

print(f"Posibles columnas de foto: {photo_cols}")

if photo_cols:
    for idx, row in matches.iterrows():
        print(f"\nIndex: {idx}")
        for pc in photo_cols:
            val = row[pc]
            print(f"  {pc}: {val}")
