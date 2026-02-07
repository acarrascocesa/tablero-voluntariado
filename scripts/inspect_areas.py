import pandas as pd

FILE_PATH = "Voluntariado Base + WPForms - Areas (Dedup Nombre) - Pais Normalizado.xlsx"
SHEET_NAME = "Merged"

try:
    try:
        df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
    except ValueError:
        print(f"Hoja '{SHEET_NAME}' no encontrada. Intentando con la primera hoja...")
        df = pd.read_excel(FILE_PATH, sheet_name=0)
    
    print("--- Columnas relacionadas con Áreas de Interés ---")
    area_cols = [c for c in df.columns if "area" in c.lower() or "área" in c.lower() or "interés" in c.lower() or "interes" in c.lower()]
    
    for col in area_cols:
        print(f"\nColumna: {col}")
        print(f"Tipo: {df[col].dtype}")
        print(f"Muestras (no nulas):")
        print(df[col].dropna().unique()[:5])
        
except Exception as e:
    print(f"Error: {e}")
