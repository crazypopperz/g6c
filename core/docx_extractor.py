"""
core/docx_extractor.py
Parse directement le synthese.docx sans Mistral.
Structure réelle : 3 tableaux par élève (index i, i+1, i+2) :
  i   : en-tête  1L×1C  → "NOM Prénom [★ BILANGUE]"
  i+1 : corps    9L×2C  → Sexe|val, Niveau|val, ..., Collège|val
  i+2 : obs      1L×1C  → "Observation : texte"
"""
from __future__ import annotations
from pathlib import Path
from docx import Document
from docx.table import Table
from docx.oxml.ns import qn
from core.models import Eleve, Sexe, Niveau, TypeAide


def _cell(table: Table, row: int, col: int) -> str:
    try:
        return table.rows[row].cells[col].text.strip()
    except IndexError:
        return ""


def _parse_niveau(v: str) -> Niveau:
    return {"TB": Niveau.TRES_BON, "B": Niveau.BON,
            "M": Niveau.MOYEN, "F": Niveau.FRAGILE}.get(v.strip().upper(), Niveau.INCONNU)


def _parse_aide(v: str) -> TypeAide:
    for a in ["PPS", "PAP", "PAI", "PPRE"]:
        if a in v.upper():
            return TypeAide(a)
    return TypeAide.AUCUNE


def extract_eleves_from_docx(path: str | Path) -> list[Eleve]:
    """Extrait les Eleve depuis un synthese.docx. Parsing Python pur, sans Mistral."""
    doc = Document(str(path))
    tables = doc.tables
    eleves = []

    # Récupérer école/commune depuis les headings du document
    ecole_courante = ""
    commune_courante = ""
    body = doc.element.body
    # Index des tableaux dans le body (pour les associer aux headings)
    tbl_elements = [c for c in body if c.tag.split("}")[-1] == "tbl"]
    tbl_index = {id(t): i for i, t in enumerate(tbl_elements)}

    # Map heading → index de tableau suivant
    heading_before: dict[int, tuple[str, str]] = {}
    ecole_tmp = ""
    commune_tmp = ""
    for child in body:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            if "—" in text and len(text) > 5:
                parts = text.split("—", 1)
                ecole_tmp   = parts[0].strip()
                commune_tmp = parts[1].strip() if len(parts) > 1 else ""
        elif tag == "tbl":
            idx = tbl_index.get(id(child))
            if idx is not None:
                heading_before[idx] = (ecole_tmp, commune_tmp)

    # Parcourir les tableaux par triplets
    i = 0
    while i < len(tables) - 1:
        t_header = tables[i]
        t_body   = tables[i + 1] if i + 1 < len(tables) else None
        t_obs    = tables[i + 2] if i + 2 < len(tables) else None

        # Vérifier structure en-tête : 1L×1C
        if len(t_header.rows) != 1 or len(t_header.rows[0].cells) != 1:
            i += 1
            continue

        nom_raw = t_header.rows[0].cells[0].text.strip()
        est_bilangue_header = "★ BILANGUE" in nom_raw
        nom_raw = nom_raw.replace("★ BILANGUE", "").strip()

        # Vérifier structure corps : 9L×2C, commence par "Sexe"
        if t_body is None or len(t_body.rows) < 9 or _cell(t_body, 0, 0) != "Sexe":
            i += 1
            continue

        # Parser le nom
        parts = nom_raw.split()
        if len(parts) >= 2:
            nom    = " ".join(parts[:-1]).upper()
            prenom = parts[-1]
        else:
            nom    = nom_raw.upper()
            prenom = ""

        if not nom:
            i += 3
            continue

        # École depuis le heading précédent
        ecole, commune = heading_before.get(i, ("", ""))

        # Observation
        obs = ""
        if t_obs and len(t_obs.rows) == 1:
            obs_raw = t_obs.rows[0].cells[0].text.strip()
            obs = obs_raw.replace("Observation : ", "").replace("Observation :", "").strip()
            # Si obs table n'est pas une obs, on ne la consomme pas
            if not obs_raw.startswith("Observation"):
                obs = ""
                i += 2  # sauter seulement header + body
            else:
                i += 3  # sauter header + body + obs
        else:
            i += 2

        # Parser le corps
        bilangue_val = _cell(t_body, 3, 1).upper() == "OUI"
        try:
            eleve = Eleve(
                nom              = nom,
                prenom           = prenom,
                sexe             = Sexe.GARCON if "arçon" in _cell(t_body, 0, 1) else Sexe.FILLE,
                ecole            = ecole,
                commune          = commune,
                niveau           = _parse_niveau(_cell(t_body, 1, 1)),
                aide             = _parse_aide(_cell(t_body, 2, 1)),
                est_bilangue     = est_bilangue_header or bilangue_val,
                est_basket       = _cell(t_body, 4, 1).upper() == "OUI",
                est_penible      = _cell(t_body, 5, 1).upper() == "OUI",
                a_separer_de     = [n.strip().upper() for n in _cell(t_body, 6, 1).split(",") if n.strip()],
                a_regrouper_avec = [n.strip().upper() for n in _cell(t_body, 7, 1).split(",") if n.strip()],
                college          = _cell(t_body, 8, 1) or "SMAC",
                observation      = obs,
            )
            eleves.append(eleve)
        except Exception as exc:
            print(f"[docx_extractor] Ignoré {nom} : {exc}")

    return eleves


def fusionner_eleves(anciens: list[Eleve], nouveaux: list[Eleve]) -> list[Eleve]:
    """Fusionne sans doublon. Les anciens ont priorité (annotations préservées)."""
    index = {e.identifiant: e for e in anciens}
    for e in nouveaux:
        if e.identifiant not in index:
            index[e.identifiant] = e
    return list(index.values())