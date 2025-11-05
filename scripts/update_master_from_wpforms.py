import argparse
import os
import re
from datetime import datetime
import pandas as pd


def norm_email(x):
    s = str(x).strip().lower()
    return s if s else None


def norm_phone(x):
    digits = re.sub(r"\D+", "", str(x))
    return digits if digits and len(digits) >= 7 else None


def norm_id(x):
    s = re.sub(r"\W+", "", str(x)).lower()
    return s if s else None


def norm_str(s):
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def norm_dob(x):
    try:
        ts = pd.to_datetime(x, dayfirst=True, errors="coerce")
        return ts.date() if pd.notna(ts) else None
    except Exception:
        return None


def pick_column(cols, candidates):
    """Regresa el nombre de columna que hace match con alguna de las expresiones en candidates."""
    lower_map = {c.lower(): c for c in cols}
    for patt in candidates:
        # exact
        if patt.lower() in lower_map:
            return lower_map[patt.lower()]
        # regex contains
        for lc, orig in lower_map.items():
            if re.search(patt, lc):
                return orig
    return None


def build_keys(df):
    cols = list(df.columns)
    email_col = pick_column(cols, [r"^correo", r"email", r"e[- ]?mail"])  # "Correo electrónico", "Email"
    phone_col = pick_column(cols, [r"tel[ée]fono", r"phone", r"cel"])
    id_col = pick_column(cols, [r"identificaci[oó]n", r"^id$"])
    full_name_col = pick_column(cols, [r"nombre completo"])  # "Nombre completo"
    first_name_col = pick_column(cols, [r"nombre completo: ?first", r"^nombre$", r"first name"])  # WPForms
    last_name_col = pick_column(cols, [r"nombre completo: ?last", r"^apellido", r"last name"])  # WPForms
    dob_col = pick_column(cols, [r"fecha de nacimiento", r"cumplea[ñn]os", r"birth"])

    emails = set()
    phones = set()
    ids = set()
    names_dob = set()

    for _, r in df.iterrows():
        e = norm_email(r.get(email_col)) if email_col else None
        p = norm_phone(r.get(phone_col)) if phone_col else None
        i = norm_id(r.get(id_col)) if id_col else None

        if full_name_col and pd.notna(r.get(full_name_col)):
            nm_base = r.get(full_name_col)
        else:
            nm_base = f"{r.get(first_name_col, '')} {r.get(last_name_col, '')}" if (first_name_col or last_name_col) else None
        nm = norm_str(nm_base) if nm_base else None

        dob = norm_dob(r.get(dob_col)) if dob_col else None

        if e:
            emails.add(e)
        if p:
            phones.add(p)
        if i:
            ids.add(i)
        if nm and dob:
            names_dob.add((nm, dob))

    return {
        "email_col": email_col,
        "phone_col": phone_col,
        "id_col": id_col,
        "full_name_col": full_name_col,
        "first_name_col": first_name_col,
        "last_name_col": last_name_col,
        "dob_col": dob_col,
        "emails": emails,
        "phones": phones,
        "ids": ids,
        "names_dob": names_dob,
    }


def is_duplicate(row, keys):
    e = norm_email(row.get(keys["email_col"])) if keys["email_col"] else None
    p = norm_phone(row.get(keys["phone_col"])) if keys["phone_col"] else None
    i = norm_id(row.get(keys["id_col"])) if keys["id_col"] else None

    if keys["full_name_col"] and pd.notna(row.get(keys["full_name_col"])):
        nm_base = row.get(keys["full_name_col"])
    else:
        nm_base = f"{row.get(keys['first_name_col'], '')} {row.get(keys['last_name_col'], '')}" if (keys['first_name_col'] or keys['last_name_col']) else None
    nm = norm_str(nm_base) if nm_base else None

    dob = norm_dob(row.get(keys["dob_col"])) if keys["dob_col"] else None

    return (
        (e and e in keys["emails"]) or
        (p and p in keys["phones"]) or
        (i and i in keys["ids"]) or
        (nm and dob and (nm, dob) in keys["names_dob"]) 
    )


# ---- Normalización de País (consistente con normalizar_pais.py) ----
PAIS_CANDIDATES = [
    "País (normalizado)",
    "País de residencia",
    "País",
    "Pais",
    "País (Residencia)",
    "Country",
]

CANON_EQUIV = {
    "usa": "Estados Unidos",
    "eeuu": "Estados Unidos",
    "estados unidos": "Estados Unidos",
    "republica dominicana": "República Dominicana",
    "mexico": "México",
    "peru": "Perú",
    "espana": "España",
    "canada": "Canadá",
}


def strip_accents_lower(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    # Normalizar acentos a ASCII
    s = (
        pd.Series([s]).str.normalize("NFKD").str.encode("ascii", "ignore").str.decode("ascii").iloc[0]
    )
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return s.strip()


def title_case_ascii(norm: str) -> str:
    return " ".join(w.capitalize() for w in norm.split())


def canonical_country(raw: str) -> str:
    norm = strip_accents_lower(raw)
    if not norm:
        return ""
    if norm in CANON_EQUIV:
        return CANON_EQUIV[norm]
    return title_case_ascii(norm)


def compute_country_norm(row, df_cols):
    for c in PAIS_CANDIDATES:
        if c in df_cols:
            val = str(row.get(c, "")).strip()
            if val:
                return canonical_country(val)
    return ""


# ---- Normalización de Áreas (consistente con normalize_areas_interes.py) ----
def is_area_column(col_name: str) -> bool:
    if col_name is None:
        return False
    s = str(col_name)
    return "Área(s) de Interés para Voluntariado" in s


def extract_area_label(col_name: str) -> str:
    s = str(col_name)
    parts = s.split(":", 1)
    label = parts[1].strip() if len(parts) > 1 else s
    label = re.sub(r"\s+", " ", label)
    return label.strip().strip('"')


def collect_areas_for_row(row, area_cols, area_labels):
    labels = []
    positives = {"si", "sí", "yes", "true", "1", "on", "checked", "selected", "x", "✓"}
    for c in area_cols:
        val = row.get(c, None)
        # Ignorar NaN/None
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        s = str(val).strip()
        # Ignorar vacíos y marcadores de nulo comunes
        if s == "" or s.lower() in {"nan", "none", "null"}:
            continue
        lab = area_labels.get(c)
        if not lab:
            continue
        # Selección válida si coincide con etiqueta o es un token positivo
        if s.lower() in positives or s.strip().lower() == str(lab).strip().lower():
            labels.append(lab)
        # de lo contrario, no contar como selección
    return labels


def main(master_path, wpforms_path, sheet_name):
    # Backup del master
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = f"{os.path.splitext(master_path)[0]}-{ts}.xlsx"
    try:
        if os.path.exists(master_path):
            import shutil
            shutil.copy2(master_path, backup_path)
            print(f"Backup creado: {backup_path}")
    except Exception as e:
        print(f"[WARN] No se pudo crear backup: {e}")

    # Lecturas
    dfm = pd.read_excel(master_path, sheet_name=sheet_name)
    xls = pd.ExcelFile(wpforms_path)
    dfn = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

    keys_m = build_keys(dfm)

    # Mapas para actualización in-place en master
    emails_map, phones_map, ids_map, namesdob_map = {}, {}, {}, {}
    for idx, r in dfm.iterrows():
        e = norm_email(r.get(keys_m["email_col"])) if keys_m["email_col"] else None
        p = norm_phone(r.get(keys_m["phone_col"])) if keys_m["phone_col"] else None
        i = norm_id(r.get(keys_m["id_col"])) if keys_m["id_col"] else None
        if keys_m["full_name_col"] and pd.notna(r.get(keys_m["full_name_col"])):
            nm_base = r.get(keys_m["full_name_col"]) 
        else:
            nm_base = f"{r.get(keys_m['first_name_col'], '')} {r.get(keys_m['last_name_col'], '')}" if (keys_m['first_name_col'] or keys_m['last_name_col']) else None
        nm = norm_str(nm_base) if nm_base else None
        dob = norm_dob(r.get(keys_m["dob_col"])) if keys_m["dob_col"] else None
        if e:
            emails_map.setdefault(e, []).append(idx)
        if p:
            phones_map.setdefault(p, []).append(idx)
        if i:
            ids_map.setdefault(i, []).append(idx)
        if nm and dob:
            namesdob_map.setdefault((nm, dob), []).append(idx)

    # Detectar columnas de áreas en WPForms y calcular normalizaciones para dfn
    area_cols = [c for c in dfn.columns if is_area_column(c)]
    area_labels = {c: extract_area_label(c) for c in area_cols}

    dup_flags = dfn.apply(lambda r: is_duplicate(r, keys_m), axis=1)
    new_only = dfn[~dup_flags].copy()

    dfn["__pais_norm__"] = dfn.apply(lambda r: compute_country_norm(r, dfn.columns), axis=1)
    dfn["__areas_list__"] = dfn.apply(lambda r: 
        "; ".join(collect_areas_for_row(r, area_cols, area_labels)) if area_cols else "",
        axis=1,
    )
    dfn["__areas_count__"] = dfn["__areas_list__"].apply(lambda s: len([x for x in str(s).split(";") if x.strip()]))

    print(f"WPForms total: {len(dfn)} | Nuevos a insertar: {len(new_only)} | Duplicados omitidos: {int(dup_flags.sum())}")

    # Actualizar in-place master para los duplicados si faltan valores
    updates = 0
    pais_col_master = "País (normalizado)" if "País (normalizado)" in dfm.columns else None
    areas_list_col = "Áreas de interés (lista)" if "Áreas de interés (lista)" in dfm.columns else None
    areas_count_col = "Áreas de interés (count)" if "Áreas de interés (count)" in dfm.columns else None

    total_labels = len(set(area_labels.values())) if area_labels else 0
    for _, r in dfn[dup_flags].iterrows():
        e = norm_email(r.get(keys_m["email_col"])) if keys_m["email_col"] else None
        p = norm_phone(r.get(keys_m["phone_col"])) if keys_m["phone_col"] else None
        i = norm_id(r.get(keys_m["id_col"])) if keys_m["id_col"] else None
        if keys_m["full_name_col"] and pd.notna(r.get(keys_m["full_name_col"])):
            nm_base = r.get(keys_m["full_name_col"]) 
        else:
            nm_base = f"{r.get(keys_m['first_name_col'], '')} {r.get(keys_m['last_name_col'], '')}" if (keys_m['first_name_col'] or keys_m['last_name_col']) else None
        nm = norm_str(nm_base) if nm_base else None
        dob = norm_dob(r.get(keys_m["dob_col"])) if keys_m["dob_col"] else None

        target_idxs = []
        if e and e in emails_map:
            target_idxs = emails_map[e]
        elif p and p in phones_map:
            target_idxs = phones_map[p]
        elif i and i in ids_map:
            target_idxs = ids_map[i]
        elif nm and dob and (nm, dob) in namesdob_map:
            target_idxs = namesdob_map[(nm, dob)]

        for ti in target_idxs:
            changed = False
            if pais_col_master:
                cur = str(dfm.loc[ti, pais_col_master]).strip() if pd.notna(dfm.loc[ti, pais_col_master]) else ""
                newv = str(r["__pais_norm__"]).strip()
                if not cur and newv:
                    dfm.loc[ti, pais_col_master] = newv
                    changed = True
            if areas_list_col:
                cur_list = str(dfm.loc[ti, areas_list_col]).strip() if pd.notna(dfm.loc[ti, areas_list_col]) else ""
                try:
                    cur_count = int(dfm.loc[ti, areas_count_col]) if (areas_count_col and pd.notna(dfm.loc[ti, areas_count_col])) else 0
                except Exception:
                    cur_count = 0
                new_list = str(r["__areas_list__"]).strip()
                new_count = int(r["__areas_count__"]) if r["__areas_count__"] else 0
                # Reglas de reparación: sobrescribir si el actual está sospechoso (todas las etiquetas) o mayor al nuevo
                suspicious_all = total_labels and cur_count == total_labels
                if (new_count > 0 and (cur_list == "" or suspicious_all or cur_count > new_count)) or (new_count == 0 and suspicious_all):
                    dfm.loc[ti, areas_list_col] = new_list
                    changed = True
                if areas_count_col and ((new_count > 0 and (cur_list == "" or suspicious_all or cur_count > new_count)) or (new_count == 0 and suspicious_all)):
                    dfm.loc[ti, areas_count_col] = new_count
                    changed = True
            if changed:
                updates += 1

    print(f"Actualizaciones in-place en master: {updates}")

    # Normalizar país/áreas de los nuevos antes de alinear columnas
    if len(new_only) > 0:
        new_only["País (normalizado)"] = new_only.apply(lambda r: compute_country_norm(r, new_only.columns), axis=1)
        if area_cols:
            new_only["Áreas de interés (lista)"] = new_only.apply(lambda r: 
                "; ".join(collect_areas_for_row(r, area_cols, area_labels)), axis=1)
            new_only["Áreas de interés (count)"] = new_only["Áreas de interés (lista)"].apply(lambda s: len([x for x in str(s).split(";") if x.strip()]))
        else:
            new_only["Áreas de interés (lista)"] = ""
            new_only["Áreas de interés (count)"] = 0

    for c in dfm.columns:
        if c not in new_only.columns:
            new_only[c] = None
    new_only = new_only[dfm.columns]

    updated = pd.concat([dfm, new_only], ignore_index=True)
    tmp_out = f"{os.path.splitext(master_path)[0]}-TEMP.xlsx"
    with pd.ExcelWriter(tmp_out, engine="openpyxl") as w:
        updated.to_excel(w, index=False, sheet_name=sheet_name)
    print(f"Escrito temporal: {tmp_out}")

    # Reemplazo atómico
    os.replace(tmp_out, master_path)
    print(f"Actualizado master: {master_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Actualiza el Excel master desde un export de WPForms con deduplicación")
    parser.add_argument("--master", required=True, help="Ruta del Excel master")
    parser.add_argument("--wpforms", required=True, help="Ruta del export de WPForms")
    parser.add_argument("--sheet", default="Merged", help="Nombre de la hoja del master")
    args = parser.parse_args()
    main(args.master, args.wpforms, args.sheet)