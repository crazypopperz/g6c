"""
core/extractor.py
Appel Mistral pour extraire les données structurées depuis le texte du file_parser.
Retourne une liste d'objets Eleve validés par Pydantic.

Modèle : mistral-small-latest
"""

from __future__ import annotations
import json
import os
from pathlib import Path
from mistralai import Mistral
from core.models import Eleve, Sexe, Niveau, TypeAide

# ─── Client ───────────────────────────────────────────────────────────────────

def _client() -> Mistral:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / ".env")
    key = os.environ.get("MISTRAL_API_KEY", "")
    if not key:
        raise EnvironmentError("MISTRAL_API_KEY non définie.")
    return Mistral(api_key=key)


# ─── Prompt système ───────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """\
Tu es un extracteur de données scolaires. Tu reçois un texte décrivant des élèves \
de CM2 issus d'un tableau de commission de liaison CM2-6e.
Le texte peut être pré-structuré (format ELEVE: ...) ou brut (extrait PDF).

Retourne UNIQUEMENT un objet JSON valide, sans markdown, sans texte avant ou après.

Format attendu :
{
  "commune": "string",
  "ecole": "string",
  "enseignant": "string",
  "eleves": [
    {
      "nom": "NOM EN MAJUSCULES",
      "prenom": "Prénom avec majuscule initiale",
      "sexe": "G" ou "F",
      "niveau": "TB" | "B" | "M" | "F" | "?",
      "aide": "" | "PAP" | "PAI" | "PPS" | "PPRE" | combinaisons séparées par virgule,
      "bilangue": true | false,
      "basket": true | false,
      "penible": true | false,
      "a_separer_de": ["NOM1", "NOM2"],
      "a_regrouper_avec": ["NOM1", "NOM2"],
      "college": "SMAC" | "nom de l'autre établissement",
      "observation": "texte brut corrigé, sans les mentions de séparation/regroupement"
    }
  ]
}

=== RÈGLE 1 — SEXE ===
Format pré-structuré : utilise directement le champ sexe "G" ou "F".
Format PDF brut : sur la ligne de données de chaque élève, tu trouveras exactement un "1" \
dans la colonne M (masculin) OU un "1" dans la colonne F (féminin). \
La structure est : "1 [collège]" pour garçon ou "[collège] 1" selon la disposition. \
Exemple : "Balthazard Camille 1 1 STE MARIE AUX CHENES" → le premier 1 = M, deuxième 1 = F \
donc ici M=1 F=1 est impossible, c'est toujours l'un OU l'autre. \
Analyse le contexte : Camille est féminin → F. \
En cas de doute, infère depuis le prénom (Gabriel → G, Camille → F, etc.).

=== RÈGLE 2 — NIVEAU SCOLAIRE ===
Calcule le niveau depuis les 5 compétences dans cet ordre : LF, LMSI, LVE, FPC, MOA.
Correspondance STRICTE — applique EXACTEMENT ces valeurs, sans exception :
- TBM (Très Bonne Maîtrise) = 4 points
- MS (Maîtrise Satisfaisante) = 3 points
- MF (Maîtrise Fragile) = 2 points
- MI (Maîtrise Insuffisante) = 1 point

IMPORTANT : MS vaut 3 points (BON), PAS 4. MF vaut 2 points (MOYEN), PAS 3.
Le comportement de l'élève N'INFLUENCE PAS le niveau scolaire — seules les 5 compétences comptent.

Calcule la moyenne des 5 valeurs :
- Moyenne ≥ 3.5 → TB
- Moyenne ≥ 2.5 → B
- Moyenne ≥ 1.5 → M
- Moyenne < 1.5 → F

RÈGLE CRITIQUE NIVEAU — plafonds obligatoires :
a) Si LF=MI ou LMSI=MI → niveau plafonné à M (même si moyenne meilleure).
b) Si l'élève a une aide (PAP, PAI, PPS, PPRE) → niveau plafonné à M.
c) Si l'observation contient "fragile" → niveau plafonné à M.

VALIDATION : croise avec l'observation pour confirmer. \
Si désaccord, garde le résultat calculé et note-le dans l'observation.

Si les 5 codes sont absents, infère depuis l'observation uniquement :
"très bon/excellent/TBM" → TB | "bon/bien/assez bien/MS" → B | \
"moyen/passable/MF" → M | "faible/très faible/en difficulté/MI" → F.

EXEMPLES DE CALCUL OBLIGATOIRES (applique exactement la même logique) :
- MS MS MS MS MS → (3+3+3+3+3)/5 = 3.0 → B
- TBM TBM TBM TBM TBM → (4+4+4+4+4)/5 = 4.0 → TB
- MI MI MI MF MI → (1+1+1+2+1)/5 = 1.2 → F (LF=MI donc max M, mais 1.2 < 1.5 donc F)
- MF MS MF MS MF → (2+3+2+3+2)/5 = 2.4 → M
- TBM MF TBM TBM TBM → (4+2+4+4+4)/5 = 3.6 → TB
- TBM TBM MS TBM TBM → (4+4+3+4+4)/5 = 3.8 → TB
- MS MF MS MS MF → (3+2+3+3+2)/5 = 2.6 → B

=== RÈGLE 3 — AIDES ===
Cherche PAP, PAI, PPS, PPRE dans toute la ligne ET dans l'observation. \
Retourne toutes celles trouvées séparées par virgule. Exemple : "PAP,PPRE".

=== RÈGLE 4 — SÉPARATIONS ET REGROUPEMENTS ===
- "A séparer de / à séparer de / séparer absolument de" → a_separer_de
- "A laisser avec / à laisser avec / laisser avec sa jumelle / à regrouper avec" → a_regrouper_avec
- Ces listes contiennent UNIQUEMENT des NOMS DE FAMILLE en majuscules.
- Supprime les mentions de classe (CM2A, CM2B…) et les prénoms.
- Supprime ces consignes du champ observation.

=== RÈGLE 5 — BASKET ===
basket = true si : observation contient "section sportive basket" ou "section basket", \
ou si la colonne ASSN est cochée (valeur 1 dans le tableau brut).

=== RÈGLE 6 — PÉNIBLE ===
penible = true si : colonne "Diff. d'appr." cochée (valeur 1), \
ou observation mentionne "problèmes de comportement", "difficultés relationnelles", \
"beaucoup d'histoires", "beaucoup de problèmes".

=== RÈGLE 7 — COLLÈGE D'AFFECTATION ===
Chaque élève a un collège d'affectation.
- Si la ligne contient "STE MARIE AUX CHENES", "SAINTE MARIE", "SMAC", "Ernest Revenu" \
ou toute variante orthographique → college = "SMAC".
- Si la colonne collège est vide ET l'observation mentionne un autre établissement \
(ex: "Jean XXIII", "IME Rettel", "inscrit à Jean XX III") → college = nom exact trouvé.
- Si rien n'est trouvé → college = "SMAC" (valeur par défaut).

=== RÈGLE 8 — OBSERVATION FINALE ===
Contient le texte original corrigé orthographiquement, SANS les consignes \
de séparation/regroupement déjà extraites, SANS les mentions de collège déjà extraites.
Conserve les informations médicales (diabète, troubles, etc.).
Ne reformule pas.

=== RÈGLE 9 — FORMAT PDF BRUT ===
Dans le PDF brut, les données d'un élève peuvent être réparties sur plusieurs lignes \
(nom/prénom sur une ligne, données sur la suivante). \
Relie toujours le nom/prénom aux données qui le suivent immédiatement. \
Ne saute aucun élève. Le total d'élèves est indiqué en bas du tableau : vérifie-le.

=== RÈGLE 10 — QUALITÉ ===
Ne crée aucun élève fictif. Ne retourne aucune clé supplémentaire. \
Si un champ est absent, utilise la valeur par défaut ("", false, [], "SMAC").
"""

_USER_PROMPT_TEMPLATE = """\
Voici les données d'un tableau de commission CM2 :

{texte}

Retourne le JSON structuré. N'oublie aucun élève (vérifie le total indiqué).
"""


# ─── Parsing du JSON retourné ─────────────────────────────────────────────────

def _parse_niveau(v: str) -> Niveau:
    mapping = {"TB": Niveau.TRES_BON, "B": Niveau.BON, "M": Niveau.MOYEN, "F": Niveau.FRAGILE}
    return mapping.get(str(v).upper(), Niveau.INCONNU)


def _parse_aide(v: str) -> TypeAide:
    priority = ["PPS", "PAP", "PAI", "PPRE"]
    aides = [a.strip().upper() for a in str(v).split(",") if a.strip()]
    for p in priority:
        if p in aides:
            return TypeAide(p)
    return TypeAide.AUCUNE


def _json_to_eleves(data: dict, source_fichier: str = "") -> list[Eleve]:
    eleves = []
    for e in data.get("eleves", []):
        try:
            eleve = Eleve(
                nom              = str(e.get("nom", "")).upper().strip(),
                prenom           = str(e.get("prenom", "")).strip(),
                sexe             = Sexe.GARCON if e.get("sexe") == "G" else Sexe.FILLE,
                ecole            = data.get("ecole", ""),
                commune          = data.get("commune", ""),
                niveau           = _parse_niveau(e.get("niveau", "?")),
                aide             = _parse_aide(e.get("aide", "")),
                est_bilangue     = bool(e.get("bilangue", False)),
                est_basket       = bool(e.get("basket", False)),
                est_penible      = bool(e.get("penible", False)),
                a_separer_de     = [n.upper().strip() for n in e.get("a_separer_de", []) if n],
                a_regrouper_avec = [n.upper().strip() for n in e.get("a_regrouper_avec", []) if n],
                college          = str(e.get("college", "SMAC")).strip() or "SMAC",
                observation      = str(e.get("observation", "")).strip(),
            )
            eleves.append(eleve)
        except Exception as exc:
            print(f"[extractor] Élève ignoré ({e.get('nom', '?')}) : {exc}")
    return eleves


# ─── API publique ──────────────────────────────────────────────────────────────

def extract_from_text(texte_structure: str, source_fichier: str = "") -> list[Eleve]:
    client = _client()
    response = client.chat.complete(
        model       = "mistral-small-latest",
        messages    = [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user",   "content": _USER_PROMPT_TEMPLATE.format(texte=texte_structure)},
        ],
        temperature = 0.0,
        max_tokens  = 4096,
    )
    raw = response.choices[0].message.content.strip()

    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    data = json.loads(raw)
    return _json_to_eleves(data, source_fichier)


def extract_from_file(path: str | Path) -> list[Eleve]:
    from core.file_parser import parse_file
    texte = parse_file(path)
    return extract_from_text(texte, source_fichier=Path(path).name)