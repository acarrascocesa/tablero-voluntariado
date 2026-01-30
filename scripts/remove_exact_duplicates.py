import pandas as pd
import os
import re
import sys
from datetime import datetime

# Configuración
MASTER_PATH = "/Users/angelcarrasco/Documents/Programacion/jcc/tablero-voluntariado/Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
BACKUP_DIR = "/Users/angelcarrasco/Documents/Programacion/jcc/tablero-voluntariado/backups"

# Columnas clave (basado en inspección previa)
COL_EMAIL = "Correo electrónico"
COL_FOTO = "Foto de perfil"
COL_ID = "Identificación"
COL_FIRST = "Nombre completo: First"
COL_LAST = "Nombre completo: Last"
COL_DOB = "Fecha de nacimiento"
COL_PHONE = "Teléfono"

def norm_email(x):
    s = str(x).strip().lower()
    return s if s and s != "nan" else None

import unicodedata

def norm_str(s):
    if pd.isna(s): return ""
    s = str(s).strip().lower()
    # Remove accents
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return re.sub(r"\s+", " ", s)

def norm_id(x):
    if pd.isna(x): return None
    s = re.sub(r"\W+", "", str(x)).lower()
    return s if s else None

def get_full_name(row):
    f = str(row.get(COL_FIRST, ""))
    l = str(row.get(COL_LAST, ""))
    if f == "nan": f = ""
    if l == "nan": l = ""
    return norm_str(f"{f} {l}")

def has_photo(row):
    val = row.get(COL_FOTO)
    if pd.isna(val): return False
    s = str(val).strip()
    return len(s) > 0 and s.lower() != "nan"

def count_filled(row):
    return row.notna().sum()

def main():
    if not os.path.exists(MASTER_PATH):
        print(f"No se encuentra el archivo maestro: {MASTER_PATH}")
        sys.exit(1)

    print(f"Leyendo master: {MASTER_PATH}")
    df = pd.read_excel(MASTER_PATH)
    original_count = len(df)
    
    # Crear columna temporal de email normalizado
    df["_norm_email"] = df[COL_EMAIL].apply(norm_email)
    
    # Agrupar por email
    groups = df.groupby("_norm_email")
    
    indices_to_drop = []
    kept_indices = []
    
    # Reporte de acciones
    deleted_log = []
    ambiguous_log = []

    print("Analizando duplicados...")
    
    # Filas que no tienen email válido se quedan tal cual (o se podrían revisar, pero asumimos dejarlas)
    no_email_indices = df[df["_norm_email"].isna()].index.tolist()
    
    # Procesar grupos
    processed_indices = set()
    
    for email, group in groups:
        if len(group) < 2:
            continue
            
        # Para este grupo de emails duplicados, buscar subgrupos de "identidad exacta"
        # Estrategia: Tomar el primer elemento y buscar sus match, luego los restantes.
        # Pero simplifiquemos: Vamos a comparar todos contra todos o agrupar por (Nombre o ID)
        
        # Lista de diccionarios con info clave para comparar
        records = []
        for idx, row in group.iterrows():
            records.append({
                "idx": idx,
                "name": get_full_name(row),
                "id": norm_id(row.get(COL_ID)),
                "has_photo": has_photo(row),
                "filled_count": count_filled(row),
                "row": row
            })
        
        # Agrupar registros que son la "misma persona"
        # Dos registros son la misma persona si:
        # 1. Tienen el mismo ID (y ID no es nulo)
        # 2. O tienen el mismo Nombre (y Nombre no es vacío)
        
        # Usaremos un enfoque de clustering simple
        clusters = []
        
        for rec in records:
            added = False
            for cluster in clusters:
                # Comprobar si 'rec' pertenece a este cluster
                match = False
                for existing in cluster:
                    # Chequeo ID
                    if rec["id"] and existing["id"] and rec["id"] == existing["id"]:
                        match = True
                        break
                    # Chequeo Nombre
                    if rec["name"] and existing["name"] and rec["name"] == existing["name"]:
                        match = True
                        break
                
                if match:
                    cluster.append(rec)
                    added = True
                    break
            
            if not added:
                clusters.append([rec])
        
        # Procesar cada cluster
        for cluster in clusters:
            if len(cluster) > 1:
                # Encontramos duplicados exactos (mismo email + (mismo ID o mismo Nombre))
                # Elegir el ganador
                
                # Criterio 1: Tiene foto
                with_photo = [r for r in cluster if r["has_photo"]]
                if with_photo:
                    candidates = with_photo
                else:
                    candidates = cluster
                
                # Criterio 2: Más datos llenos
                candidates.sort(key=lambda x: x["filled_count"], reverse=True)
                winner = candidates[0]
                
                # Marcar perdedores para borrar
                for loser in cluster:
                    if loser["idx"] != winner["idx"]:
                        indices_to_drop.append(loser["idx"])
                        deleted_log.append({
                            "email": email,
                            "kept_idx": winner["idx"],
                            "dropped_idx": loser["idx"],
                            "reason": "Duplicado exacto (mismo nombre/ID)"
                        })
            else:
                # Este cluster tiene tamaño 1, significa que dentro del grupo de emails, 
                # este registro es único en identidad (ej. email compartido por 2 personas distintas)
                # O es el único remanente.
                # Si había otros clusters en este mismo grupo de email, significa conflicto de identidad.
                pass
        
        # Chequear si hubo clusters distintos para el mismo email (caso ambiguo)
        if len(clusters) > 1:
            # Ejemplo: mismo email, pero un cluster es "Juan" y otro es "Pedro"
            names = [c[0]["name"] for c in clusters]
            ambiguous_log.append({
                "email": email,
                "names_found": names,
                "count": len(group)
            })

    # Ejecutar borrado
    if not indices_to_drop:
        print("No se encontraron duplicados exactos para eliminar según los criterios.")
    else:
        print(f"Se encontraron {len(indices_to_drop)} registros para eliminar.")
        
        # Backup
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_file = os.path.join(BACKUP_DIR, f"backup_before_dedup_{ts}.xlsx")
        import shutil
        shutil.copy2(MASTER_PATH, backup_file)
        print(f"Backup creado: {backup_file}")
        
        # Eliminar
        df_clean = df.drop(indices_to_drop)
        # Limpiar columna temporal
        df_clean.drop(columns=["_norm_email"], inplace=True)
        
        # Guardar
        df_clean.to_excel(MASTER_PATH, index=False)
        print(f"Archivo maestro actualizado. Registros: {original_count} -> {len(df_clean)}")
        
        # Generar reporte simple de ambiguos si los hay
        if ambiguous_log:
            print("\n[INFO] Se detectaron emails compartidos por personas con identidad distinta (no eliminados):")
            for item in ambiguous_log:
                print(f"  - {item['email']}: {item['names_found']}")

if __name__ == "__main__":
    main()
