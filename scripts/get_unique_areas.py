import pandas as pd

PATH = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
SHEET = "Merged"
COL_NAME = "Áreas de interés (lista)"

def main():
    try:
        df = pd.read_excel(PATH, sheet_name=SHEET)
    except Exception as e:
        print(f"Error leyendo el archivo: {e}")
        return

    if COL_NAME not in df.columns:
        print(f"La columna '{COL_NAME}' no existe en el dataset.")
        # Intentar buscar columnas raw si la consolidada no existe
        cols = [c for c in df.columns if "área" in str(c).lower() and "interés" in str(c).lower()]
        print(f"Columnas alternativas encontradas: {cols}")
        return

    # Extraer y separar las áreas
    all_areas = []
    # Filtrar nulos y vacíos
    series = df[COL_NAME].dropna().astype(str)
    
    for val in series:
        if not val.strip():
            continue
        # Separar por punto y coma y limpiar espacios
        items = [x.strip() for x in val.split(";")]
        all_areas.extend(items)

    # Contar frecuencias
    from collections import Counter
    counts = Counter(all_areas)
    
    # Ordenar por frecuencia descendente
    sorted_areas = sorted(counts.items(), key=lambda x: x[1], reverse=True)

    print(f"Total de registros con áreas: {len(series)}")
    print(f"Áreas únicas encontradas: {len(sorted_areas)}")
    print("-" * 40)
    for area, count in sorted_areas:
        print(f"{area} ({count})")

if __name__ == "__main__":
    main()
