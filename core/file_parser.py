"""
core/file_parser.py
Lecture des tableaux de commission (xlsx, xls, ods, pdf).
Produit un texte structuré propre prêt à envoyer à Mistral.
Structure du template national (identique xlsx / ods) :
Ligne 6  → col 20 : label "Commune"  | col 22 : valeur
Ligne 7  → col 20 : label "Ecole"    | col 22 : valeur
Ligne 8  → col 20 : label "Enseignant"| col 22 : valeur
Ligne 20 → headers finaux (M, F, ANG, BI L, LF, LMSI, LVE, FPC, MOA…)
Ligne 21+→ données élèves
Mapping colonnes (0-indexé) :
1  = Nom Prénom
2  = M (garçon si 1)
3  = F (fille si 1)
10 = Diff. d'appr. (pénible)
11 = PPRE
12 = PAI
13 = PAP
14 = PPS
17 = ANG (anglais LV1)
18 = BI L (bilangue)
19 = LF   (niveau français oral/écrit)
20 = LMSI (niveau maths/sciences)
21 = LVE  (niveau langues vivantes)
22 = FPC  (formation personne/citoyen)
23 = MOA  (méthodes/outils)
24 = ASSN / Section basket (si 1)
31 = OBSERVATIONS
"""
from __future__ import annotations
import re
import pandas as pd
from pathlib import Path

# ─── Constantes de mapping ─────────────────────────────────────────────────────
COL = {
    "nom_prenom": 1,
    "M":          2,
    "F":          3,
    "penible":    10,
    "PPRE":       11,
    "PAI":        12,
    "PAP":        13,
    "PPS":        14,
    "ANG":        17,
    "bilangue":   18,
    "LF":         19,
    "LMSI":       20,
    "LVE":        21,
    "FPC":        22,
    "MOA":        23,
    "basket":     24,  # NE PAS UTILISER — col 24 = ASSN (attestation sauvetage).
                       # Basket détecté uniquement dans l'observation.
    "obs":        31,
}

_NOISE_EXACT = {
    "total", "lf", "lmsi", "lve", "fpc", "moa",
    "mi", "mf", "ms", "tbm", "tb",
    "college", "sélectionner dans menu",
    "ars p. de rozier", "marly j. mermoz", "metz g. de la tour",
    "metz site de montigny", "metz barbot",
    "moulins a. camus", "moulins l. armand",
}

_NIVEAU_MAP = {
    "TBM": "TB", "TB": "TB",
    "MS":  "B",  "B":  "B",
    "MF":  "M",  "M":  "M",
    "MI":  "F",  "F":  "F",
}

def _is_noise(nom: str) -> bool:
    n = nom.strip().lower()
    if len(n) < 3:
        return True
    return n in _NOISE_EXACT

def _niveau_global(row: pd.Series) -> str:
    """Calcule un niveau global à partir des 5 compétences LF/LMSI/LVE/FPC/MOA."""
    scores = {"TB": 4, "B": 3, "M": 2, "F": 1}
    vals = []
    for col_key in ("LF", "LMSI", "LVE", "FPC", "MOA"):
        raw = str(row[COL[col_key]]).strip().upper() if pd.notna(row[COL[col_key]]) else ""
        mapped = _NIVEAU_MAP.get(raw)
        if mapped:
            vals.append(scores[mapped])
    if not vals:
        return "?"
    avg = sum(vals) / len(vals)
    if avg >= 3.5:
        return "TB"
    elif avg >= 2.5:
        return "B"
    elif avg >= 1.5:
        return "M"
    else:
        return "F"

def _aides(row: pd.Series) -> str:
    aides = []
    for aide in ("PPRE", "PAI", "PAP", "PPS"):
        v = str(row[COL[aide]]).strip() if pd.notna(row[COL[aide]]) else ""
        if v not in ("", "nan", "0", "None"):
            aides.append(aide)
    return ", ".join(aides)

def _val(df: pd.DataFrame, row: int, col: int) -> str:
    """Lit une cellule et retourne une chaîne propre, '' si vide."""
    try:
        v = str(df.iloc[row, col]).strip()
        return "" if v.lower() in ("nan", "none") else v
    except Exception:
        return ""

def _meta_from_df(df: pd.DataFrame) -> dict:
    """
    Extrait commune, école, enseignant depuis l'en-tête du tableau.
    Scan dynamique sur les 15 premières lignes : robuste aux variations
    de template (lignes vides en tête, version de fichier différente).
    """
    meta = {"commune": "", "ecole": "", "enseignant": ""}
    targets = {
        "commune":    ["commune"],
        "ecole":      ["ecole", "école"],
        "enseignant": ["enseignant"],
    }
    for row_idx in range(min(15, len(df))):
        for col_idx in range(df.shape[1]):
            cell = _val(df, row_idx, col_idx).lower().strip()
            for key, labels in targets.items():
                if meta[key]:          # déjà trouvé, on passe
                    continue
                if cell in labels:
                    # La valeur est dans la même ligne, cherche la première
                    # cellule non vide à droite du label
                    for offset in range(1, min(6, df.shape[1] - col_idx)):
                        val = _val(df, row_idx, col_idx + offset)
                        if val and val.lower() not in labels:
                            meta[key] = val
                            break
    return meta

def _split_nom_prenom(nom_raw: str) -> tuple[str, str]:
    """
    Découpe 'NOM PRENOM' ou 'Nom Prenom' en (nom, prenom).
    Cas 1 : tout majuscules (PDF standard) → dernier mot = prénom
    Cas 2 : casse mixte (PDF casse mixte, xlsx/ods) → dernier mot = prénom
    """
    parts = nom_raw.strip().split()
    if len(parts) == 1:
        return parts[0].upper(), ""
    
    # Cas tout-majuscules : "BALTHAZARD CAMILLE"
    if all(p == p.upper() for p in parts):
        nom    = " ".join(parts[:-1]).upper()
        prenom = parts[-1].capitalize()
        return nom, prenom

    # Cas casse mixte : "Balthazard Camille" ou "DUPONT Marie"
    # Dernier mot = prénom, tout le reste = nom
    nom    = " ".join(parts[:-1]).upper()
    prenom = parts[-1].capitalize()
    return nom, prenom

def _rows_from_df(df: pd.DataFrame) -> list[dict]:
    """Extrait la liste des élèves depuis un DataFrame normalisé.
    S'arrête à la ligne 'Total' qui marque la fin du tableau élèves."""
    eleves = []
    for _, row in df.iloc[21:].iterrows():
        nom_raw = row[COL["nom_prenom"]]
        # Arrêt propre à la ligne Total
        if isinstance(nom_raw, str) and nom_raw.strip().lower() == "total":
            break
        if not isinstance(nom_raw, str) or _is_noise(nom_raw):
            continue

        nom, prenom = _split_nom_prenom(nom_raw)
        if not nom:
            continue

        sexe     = "G" if (pd.notna(row[COL["M"]]) and str(row[COL["M"]]).strip() == "1") else "F"
        
        # CORRECTION CRITIQUE : Bilangue = "BIL" ou "1" (PAS "ANG" ou "ALL" !)
        val_bil = str(row[COL["bilangue"]]).strip().upper() if pd.notna(row[COL["bilangue"]]) else ""
        bilangue = val_bil in ("BIL", "1")
        
        # CORRECTION CRITIQUE : col 24 = ASSN (attestation sauvetage), PAS basket.
        # Basket détecté UNIQUEMENT dans l'observation, comme pour le PDF.
        obs_raw  = str(row[COL["obs"]]).strip() if pd.notna(row[COL["obs"]]) else ""
        obs      = obs_raw if obs_raw.lower() not in ("nan", "none", "") else ""
        basket   = "basket" in obs.lower()
        
        penible  = pd.notna(row[COL["penible"]]) and str(row[COL["penible"]]).strip() not in ("", "nan", "0", "None")

        eleves.append({
            "nom":      nom,
            "prenom":   prenom,
            "sexe":     sexe,
            "niveau":   _niveau_global(row),
            "aide":     _aides(row),
            "bilangue": bilangue,
            "basket":   basket,
            "penible":  penible,
            "obs":      obs,
        })
    return eleves

def _df_to_text(meta: dict, eleves: list[dict]) -> str:
    """Sérialise en texte structuré pour Mistral."""
    lines = [
        f"COMMUNE: {meta['commune']}",
        f"ECOLE: {meta['ecole']}",
        f"ENSEIGNANT: {meta['enseignant']}",
        f"TOTAL_ELEVES: {len(eleves)}",
        "---",
    ]
    for e in eleves:
        aide_str     = f" | AIDE:{e['aide']}"   if e["aide"]     else ""
        bilangue_str = " | BILANGUE:oui"        if e["bilangue"] else ""
        basket_str   = " | BASKET:oui"          if e["basket"]   else ""
        penible_str  = " | PENIBLE:oui"         if e["penible"]  else ""
        obs_str      = f" | OBS:{e['obs']}"     if e["obs"]      else ""
        lines.append(
            f"ELEVE: {e['nom']} | {e['prenom']} | {e['sexe']} | "
            f"NIVEAU:{e['niveau']}"
            f"{aide_str}{bilangue_str}{basket_str}{penible_str}{obs_str}"
        )
    return "\n".join(lines)

# ── API publique ──────────────────────────────────────────────────────────────

def parse_file(path: str | Path) -> str:
    """
    Lit un tableau de commission (xlsx, xls, ods, pdf)
    et retourne un texte structuré prêt pour Mistral.
    Lève ValueError si le format n'est pas supporté.
    """
    path = Path(path)
    ext  = path.suffix.lower()
    
    if ext in (".xlsx", ".xls"):
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        df     = pd.read_excel(path, header=None, engine=engine)
        meta   = _meta_from_df(df)
        eleves = _rows_from_df(df)
        return _df_to_text(meta, eleves)

    elif ext == ".ods":
        df     = pd.read_excel(path, header=None, engine="odf")
        meta   = _meta_from_df(df)
        eleves = _rows_from_df(df)
        return _df_to_text(meta, eleves)

    elif ext == ".pdf":
        import pdfplumber

        # Mapping colonnes PDF (extract_table, 0-indexé, 33 cols)
        # Col 0  = Nom Prénom
        # Col 1  = M (garçon si "1")
        # Col 2  = F (fille si "1")  — attention : col 3 aussi présent
        # Col 6  = Collège d'affectation
        # Col 9  = Diff. d'appr. (pénible)
        # Col 11 = PPRE, 12 = PAI, 13 = PAP, 14 = PPS
        # Col 16 = Bilangue
        # Col 19 = LF, 20 = LMSI, 21 = LVE, 22 = FPC, 23 = MOA
        # Col 24 = ASSN/Basket → IGNORÉ ! On utilise UNIQUEMENT les observations
        # Col 31 = Observations
        COL_PDF = {
            "nom_prenom": 0,
            "M":          2,
            "F":          3,
            "college":    6,
            "penible":    10,
            "PPRE":       11,
            "PAI":        12,
            "PAP":        13,
            "PPS":        14,
            "bilangue":   17,
            "LF":         19,
            "LMSI":       20,
            "LVE":        21,
            "FPC":        22,
            "MOA":        23,
            "basket":     24,  # IGNORÉ - Colonne ASSN, pas basket !
            "obs":        31,
        }

        _SMAC_VARIANTS = {
            "ste marie aux chenes", "sainte marie aux chenes",
            "sainte-marie-aux-chênes", "smac", "ernest revenu",
            "ste marie", "sainte marie",
        }

        def _is_smac(val: str) -> bool:
            v = val.lower().strip()
            return any(s in v for s in _SMAC_VARIANTS)

        def _cell(row: list, idx: int) -> str:
            try:
                v = row[idx]
                return str(v).strip() if v is not None else ""
            except IndexError:
                return ""

        def _is_eleve_row(row: list) -> bool:
            """Ligne élève = col 0 contient un nom ET col 19-23 contient un code niveau."""
            nom = _cell(row, 0)
            if not nom or len(nom) < 3:
                return False
            if _is_noise(nom):
                return False
            # Au moins une des 5 compétences est renseignée
            niveaux = {_cell(row, i).upper() for i in (19, 20, 21, 22, 23)}
            return bool(niveaux & {"TBM", "MS", "MF", "MI"})

        meta = {"commune": "", "ecole": "", "enseignant": ""}
        eleve_rows = []

        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    # Chercher meta dans les lignes d'en-tête
                    for row in table:
                        if not row:
                            continue
                        for i, cell in enumerate(row):
                            if not cell:
                                continue
                            cs = str(cell).strip()
                            if "Commune" in cs and not meta["commune"]:
                                val = cs.replace("Commune", "").strip()
                                meta["commune"] = val or _cell(row, 22)
                            if cs == "Ecole" and not meta["ecole"]:
                                meta["ecole"] = _cell(row, 22)
                            if "Enseignant" in cs and not meta["enseignant"]:
                                val = cs.replace("Enseignant", "").strip()
                                meta["enseignant"] = val or _cell(row, 22)
                        # Collecter les lignes élèves
                        if _is_eleve_row(row):
                            eleve_rows.append(row)

        # Parser chaque ligne élève
        eleves = []
        for row in eleve_rows:
            try:
                nom_raw = _cell(row, COL_PDF["nom_prenom"])
                if _is_noise(nom_raw):
                    continue
                nom, prenom = _split_nom_prenom(nom_raw)
                if not nom:
                    continue
                sexe = "G" if _cell(row, COL_PDF["M"]) == "1" else "F"
                niveau_row = pd.Series({
                    COL["LF"]:   _cell(row, COL_PDF["LF"])   or None,
                    COL["LMSI"]: _cell(row, COL_PDF["LMSI"]) or None,
                    COL["LVE"]:  _cell(row, COL_PDF["LVE"])  or None,
                    COL["FPC"]:  _cell(row, COL_PDF["FPC"])  or None,
                    COL["MOA"]:  _cell(row, COL_PDF["MOA"])  or None,
                })
                niveau = _niveau_global(niveau_row)
                aides = []
                for aide in ("PPRE", "PAI", "PAP", "PPS"):
                    v = _cell(row, COL_PDF[aide])
                    if v and v not in ("", "0", "nan"):
                        aides.append(aide)
                aide_str = ", ".join(aides)
                
                # CORRECTION CRITIQUE PDF : Bilangue = "BIL" ou "1"
                val_bil = _cell(row, COL_PDF["bilangue"]).upper()
                bilangue = val_bil in ("BIL", "1")
                
                # CORRECTION CRITIQUE PDF : Basket = UNIQUEMENT via observation (col 24 = ASSN)
                basket = False
                
                penible  = _cell(row, COL_PDF["penible"]) == "1"
                college_raw = _cell(row, COL_PDF["college"])
                college = "SMAC" if (not college_raw or _is_smac(college_raw)) else college_raw
                obs = _cell(row, COL_PDF["obs"]).replace("\n", " ").strip()
                
                # Basket UNIQUEMENT depuis observation
                if "basket" in obs.lower():
                    basket = True

                if not penible and any(
                    kw in obs.lower() for kw in
                    ("problèmes de comportement", "difficultés relationnelles",
                     "beaucoup d'histoires", "beaucoup de problèmes")
                ):
                    penible = True
                eleves.append({
                    "nom": nom, "prenom": prenom, "sexe": sexe,
                    "niveau": niveau, "aide": aide_str,
                    "bilangue": bilangue, "basket": basket,
                    "penible": penible, "college": college, "obs": obs,
                })
            except Exception as exc:
                print(f"[DEBUG] Élève ignoré ({_cell(row, 0)}) : {exc}")

        # Utiliser _df_to_text en ajoutant le champ college
        lines = [
            f"COMMUNE: {meta['commune']}",
            f"ECOLE: {meta['ecole']}",
            f"ENSEIGNANT: {meta['enseignant']}",
            f"TOTAL_ELEVES: {len(eleves)}",
            "---",
        ]
        for e in eleves:
            aide_s     = f" | AIDE:{e['aide']}"     if e["aide"]     else ""
            bil_s      = " | BILANGUE:oui"          if e["bilangue"] else ""
            bask_s     = " | BASKET:oui"            if e["basket"]   else ""
            pen_s      = " | PENIBLE:oui"           if e["penible"]  else ""
            college_s  = f" | COLLEGE:{e['college']}"
            obs_s      = f" | OBS:{e['obs']}"       if e["obs"]      else ""
            lines.append(
                f"ELEVE: {e['nom']} | {e['prenom']} | {e['sexe']} | "
                f"NIVEAU:{e['niveau']}{aide_s}{bil_s}{bask_s}{pen_s}{college_s}{obs_s}"
            )
        return "\n".join(lines)

    else:
        raise ValueError(f"Format non supporté : {ext}")