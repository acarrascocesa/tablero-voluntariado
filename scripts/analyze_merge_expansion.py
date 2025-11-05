from collections import Counter
import pandas as pd


def fullname_key(df: pd.DataFrame) -> pd.Series:
    first = df['Nombre completo: First'].fillna('').astype(str).str.strip()
    last = df['Nombre completo: Last'].fillna('').astype(str).str.strip()
    return (first + ' ' + last).str.lower()


def main():
    df_base = pd.read_excel('Voluntariado Filtro - All Data.xlsx', sheet_name=0)
    df_wp = pd.read_csv('voluntarios.csv')

    base_key = fullname_key(df_base)
    wp_key = fullname_key(df_wp)

    base_counts = Counter(base_key)
    wp_counts = Counter(wp_key)

    expanded = []
    zero_matches = 0
    one_to_one = 0
    one_to_many = 0
    many_to_one = 0
    many_to_many = 0
    merged_rows = 0

    for k, b in base_counts.items():
        w = wp_counts.get(k, 0)
        if w == 0:
            zero_matches += b
            merged_rows += b
        elif b == 1 and w == 1:
            one_to_one += 1
            merged_rows += 1
        elif b == 1 and w > 1:
            one_to_many += 1
            merged_rows += w
            expanded.append((k, b, w))
        elif b > 1 and w == 1:
            many_to_one += 1
            merged_rows += b
        else:
            many_to_many += 1
            merged_rows += b * w
            expanded.append((k, b, w))

    print('Base rows:', len(df_base))
    print('WP rows:', len(df_wp))
    print('Computed merged rows:', merged_rows)
    print('Zero matches (base rows without WPForms):', zero_matches)
    print('One-to-one keys:', one_to_one)
    print('One-to-many keys:', one_to_many)
    print('Many-to-one keys:', many_to_one)
    print('Many-to-many keys:', many_to_many)
    print('\nTop expansion keys (name, base_count, wp_count)')
    for k,b,w in sorted(expanded, key=lambda x: (x[2], x[1]), reverse=True)[:15]:
        print('-', k, b, w)


if __name__ == '__main__':
    main()