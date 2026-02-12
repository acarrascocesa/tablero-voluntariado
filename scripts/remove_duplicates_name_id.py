import pandas as pd
import re
from datetime import datetime
import os

def normalize_id(x):
    s = re.sub(r"\W+", "", str(x)).lower()
    if s in ["cedula", "pasaporte", "nan", "none", "n/a", "no", "0", "porasignar"]:
        return None
    return s if s and len(s) > 4 else None

def normalize_str(s):
    if pd.isna(s): return ""
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def remove_duplicates_name_id(master_path):
    print(f"Leyendo maestro: {master_path}...")
    df = pd.read_excel(master_path)
    original_count = len(df)
    
    # Identificar columnas
    cols = df.columns
    id_col = next((c for c in cols if "identi" in c.lower() or "cedula" in c.lower()), None)
    name_cols = [c for c in cols if "nombre" in c.lower()]
    full_name_col = next((c for c in name_cols if "completo" in c.lower() and "first" not in c.lower() and "last" not in c.lower()), None)
    first_col = next((c for c in name_cols if "first" in c.lower()), None)
    last_col = next((c for c in name_cols if "last" in c.lower()), None)
    date_col = next((c for c in cols if "date" in c.lower() or "fecha" in c.lower() or "entry" in c.lower()), None)

    # Crear llaves de normalización
    print("Normalizando datos para identificación de duplicados...")
    
    def get_full_name(row):
        if full_name_col and pd.notna(row[full_name_col]):
            return normalize_str(row[full_name_col])
        elif first_col and last_col:
            return normalize_str(f"{row[first_col]} {row[last_col]}")
        return ""

    df['_key_name'] = df.apply(get_full_name, axis=1)
    df['_key_id'] = df[id_col].apply(normalize_id) if id_col else None
    
    # Crear llave combinada
    df['_key_name_id'] = df.apply(lambda r: f"{r['_key_name']}|{r['_key_id']}" if r['_key_name'] and r['_key_id'] else None, axis=1)
    
    # Calcular "completitud" del registro (cuántos campos no nulos tiene)
    df['_completeness'] = df.notna().sum(axis=1)
    
    # Ordenar por:
    # 1. Llave combinada
    # 2. Completitud (más campos llenos primero)
    # 3. Fecha (más reciente primero si existe)
    sort_cols = ['_key_name_id', '_completeness']
    ascending = [True, False]
    
    if date_col:
        df['_sort_date'] = pd.to_datetime(df[date_col], errors='coerce')
        sort_cols.append('_sort_date')
        ascending.append(False)
    
    print("Ordenando y eliminando duplicados...")
    # Solo aplicar a los que tienen llave válida
    is_duplicate_candidate = df['_key_name_id'].notna()
    
    df_valid = df[is_duplicate_candidate].sort_values(by=sort_cols, ascending=ascending)
    df_valid_clean = df_valid.drop_duplicates(subset=['_key_name_id'], keep='first')
    
    # Combinar con los registros que no eran candidatos a deduplicación por falta de ID o Nombre
    df_others = df[~is_duplicate_candidate]
    df_final = pd.concat([df_valid_clean, df_others], ignore_index=True)
    
    # Limpiar columnas temporales
    temp_cols = [c for c in df_final.columns if c.startswith('_')]
    df_final = df_final.drop(columns=temp_cols)
    
    # Backup
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_name = f"backups/{os.path.basename(master_path).replace('.xlsx', '')}_PRE_CLEANUP_{timestamp}.xlsx"
    if not os.path.exists('backups'): os.makedirs('backups')
    
    print(f"Creando backup en: {backup_name}")
    df.drop(columns=temp_cols).to_excel(backup_name, index=False)
    
    # Guardar maestro
    print(f"Guardando maestro actualizado: {master_path}")
    df_final.to_excel(master_path, index=False)
    
    final_count = len(df_final)
    print(f"\n--- RESULTADOS DE LA LIMPIEZA ---")
    print(f"Filas originales: {original_count}")
    print(f"Filas eliminadas: {original_count - final_count}")
    print(f"Filas finales: {final_count}")
    print(f"----------------------------------")

if __name__ == "__main__":
    MASTER = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
    remove_duplicates_name_id(MASTER)
