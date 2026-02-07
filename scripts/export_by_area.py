import pandas as pd
import re
import os

INPUT_FILE = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
OUTPUT_FILE = "Voluntarios_Por_Area.xlsx"
SHEET_NAME = "Merged"
AREAS_COL = "Áreas de interés (lista)"

def clean_sheet_name(name):
    """Clean string to be valid Excel sheet name (max 31 chars, no special chars)"""
    # Remove invalid characters: : \ / ? * [ ]
    clean = re.sub(r'[\\/*?:\[\]]', '', str(name))
    # Shorten to 31 chars
    return clean[:31]

def main():
    print(f"Leyendo {INPUT_FILE}...")
    try:
        try:
            df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
        except ValueError:
            df = pd.read_excel(INPUT_FILE, sheet_name=0)
    except Exception as e:
        print(f"Error leyendo el archivo: {e}")
        return

    # Normalizar columna de aréas
    if AREAS_COL not in df.columns:
        print(f"Error: Columna '{AREAS_COL}' no encontrada.")
        return

    # Obtener todas las áreas únicas
    all_areas = set()
    rows_with_area = set()
    
    # Pre-calcular sets de índices por área para velocidad
    area_to_indices = {}

    for idx, row in df.iterrows():
        areas_str = str(row[AREAS_COL])
        if pd.isna(row[AREAS_COL]) or areas_str.strip() == "" or areas_str.lower() == "nan":
            continue
            
        # Separar por punto y coma
        items = [x.strip() for x in areas_str.split(";") if x.strip()]
        
        if items:
            rows_with_area.add(idx)
            for area in items:
                all_areas.add(area)
                if area not in area_to_indices:
                    area_to_indices[area] = []
                area_to_indices[area].append(idx)

    print(f"Encontradas {len(all_areas)} áreas únicas.")
    
    # Escribir a Excel
    print(f"Escribiendo a {OUTPUT_FILE}...")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # 1. Hoja por cada área
        for area in sorted(list(all_areas)):
            indices = area_to_indices.get(area, [])
            if indices:
                subset = df.loc[indices]
                sheet_name = clean_sheet_name(area)
                print(f" - Hoja '{sheet_name}': {len(subset)} voluntarios")
                subset.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 2. Hoja para los que no tienen área
        all_indices = set(df.index)
        no_area_indices = list(all_indices - rows_with_area)
        if no_area_indices:
            no_area_df = df.loc[no_area_indices]
            print(f" - Hoja 'Sin Área Asignada': {len(no_area_df)} voluntarios")
            no_area_df.to_excel(writer, sheet_name="Sin Área Asignada", index=False)
        else:
            print(" - Todos los voluntarios tienen al menos un área.")

    print(f"\nProceso completado. Archivo guardado en: {os.path.abspath(OUTPUT_FILE)}")

if __name__ == "__main__":
    main()
