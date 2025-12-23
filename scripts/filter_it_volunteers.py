import pandas as pd
import re

PATH = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
SHEET = "Merged"
COL_NAME = "Áreas de interés (lista)"

# Definir las áreas exactas de TI según el dataset
IT_AREAS = [
    "Tecnología Sedes (Conectividad y Redes)",
    "Tecnología Deportiva",
    "Radios, Impresoras y Pantallas"
]

def main():
    try:
        df = pd.read_excel(PATH, sheet_name=SHEET)
    except Exception as e:
        print(f"Error leyendo el archivo: {e}")
        return

    if COL_NAME not in df.columns:
        print(f"La columna '{COL_NAME}' no existe.")
        return

    # Función para chequear si el voluntario seleccionó alguna de las áreas IT
    def has_it_area(val):
        if not isinstance(val, str):
            return False
        user_areas = [x.strip() for x in val.split(";")]
        # Intersección entre áreas del usuario y áreas IT
        match = set(user_areas).intersection(set(IT_AREAS))
        return len(match) > 0

    # Filtrar
    mask = df[COL_NAME].apply(has_it_area)
    df_it = df[mask].copy()

    # Seleccionar columnas relevantes para el reporte
    cols_export = [
        "Nombre completo", 
        "Correo electrónico", 
        "Teléfono", 
        "País (normalizado)", 
        "Áreas de interés (lista)"
    ]
    
    # Asegurar que las columnas existan, si no, usar fallback
    final_cols = [c for c in cols_export if c in df_it.columns]
    
    # Agregar columnas si faltan usando nombres probables
    if "Nombre completo" not in final_cols:
        if "Nombre completo: First" in df_it.columns:
            df_it["Nombre completo"] = df_it["Nombre completo: First"] + " " + df_it["Nombre completo: Last"]
            final_cols.insert(0, "Nombre completo")

    df_export = df_it[final_cols]

    # Guardar reporte
    output_path = "Reporte_Voluntarios_TI.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="TI")

    print(f"Áreas TI filtradas: {IT_AREAS}")
    print(f"Total voluntarios encontrados: {len(df_export)}")
    print(f"Reporte generado: {output_path}")
    
    # Mostrar primeros 5
    print("\nEjemplo de primeros 5 voluntarios:")
    print(df_export.head().to_string(index=False))

if __name__ == "__main__":
    main()
