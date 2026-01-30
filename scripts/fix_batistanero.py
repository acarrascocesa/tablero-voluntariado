
import pandas as pd
import shutil
from datetime import datetime
import os

master_path = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
target_email = "batistanero7@gmail.com"

# Backup
ts = datetime.now().strftime("%Y%m%d-%H%M%S")
backup_path = f"{os.path.splitext(master_path)[0]}-{ts}.xlsx"
shutil.copy2(master_path, backup_path)
print(f"Backup creado: {backup_path}")

# Load
df = pd.read_excel(master_path)
print(f"Total filas antes: {len(df)}")

# Identify rows to drop
# We know from inspection that indices 4604, 4605, 4606 (0-based in current load) are the ones without photo
# But let's be robust and find them dynamically again
email_cols = [c for c in df.columns if "correo" in c.lower() or "email" in c.lower()]
email_col = email_cols[0]
photo_col = "Foto de perfil"

# Find all rows for this email
matches = df[df[email_col].astype(str).str.strip().str.lower() == target_email]

# Split into keep and drop
to_keep = []
to_drop = []

for idx, row in matches.iterrows():
    photo_val = str(row[photo_col])
    if photo_val and photo_val.lower() != "nan" and "http" in photo_val:
        to_keep.append(idx)
    else:
        to_drop.append(idx)

# Safety check
if len(to_keep) == 1:
    print(f"Conservando fila {to_keep[0]} con foto.")
    print(f"Eliminando filas {to_drop} sin foto.")
    
    df_clean = df.drop(to_drop)
    print(f"Total filas después: {len(df_clean)}")
    
    # Save
    df_clean.to_excel(master_path, index=False)
    print("Master actualizado correctamente.")
else:
    print(f"ERROR: Se encontraron {len(to_keep)} filas con foto (se esperaba 1) y {len(to_drop)} sin foto.")
    print("No se realizaron cambios.")
