import sys
import pandas as pd


def fullname_key(df: pd.DataFrame) -> pd.Series:
    first = df['Nombre completo: First'].fillna('').astype(str).str.strip()
    last = df['Nombre completo: Last'].fillna('').astype(str).str.strip()
    return (first + ' ' + last).str.lower()


def nonempty_count(row: pd.Series) -> int:
    cnt = 0
    for val in row.values:
        if pd.isna(val):
            continue
        s = str(val).strip()
        if s == '' or s.lower() == 'nan':
            continue
        cnt += 1
    return cnt


def main(input_xlsx: str, output_xlsx: str) -> None:
    df = pd.read_excel(input_xlsx, sheet_name='Merged')
    before = len(df)

    key = fullname_key(df)
    df['__key__'] = key
    df['__fill__'] = df.apply(nonempty_count, axis=1)

    # For cada nombre, elegir la fila con mayor cantidad de campos no vacíos.
    # Si hay empate, pandas idxmax devuelve la primera aparición.
    idx_best = df.groupby('__key__')['__fill__'].idxmax()
    df_best = df.loc[idx_best].copy()

    # Ordenar por el orden original para que sea estable
    df_best = df_best.sort_index()

    # Limpiar columnas técnicas
    df_best = df_best.drop(columns=['__key__', '__fill__'])

    after = len(df_best)

    with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
        df_best.to_excel(writer, index=False, sheet_name='Merged')

    print('Input:', input_xlsx)
    print('Output:', output_xlsx)
    print('Filas antes:', before)
    print('Filas después (dedup nombre):', after)
    print('Duplicados eliminados:', before - after)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Uso: .venv/bin/python dedupe_por_nombre.py 'Voluntariado Base + WPForms - Areas.xlsx' 'Voluntariado Base + WPForms - Areas (Dedup Nombre).xlsx'")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])