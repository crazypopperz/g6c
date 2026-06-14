"""
core/models.py
Entités de données du projet G6C.
Toutes les structures qui traversent l'application sont définies ici.
"""

from __future__ import annotations
from enum import Enum
from typing import Optional
from pydantic import BaseModel, field_validator


# ─────────────────────────────────────────
#  Enums
# ─────────────────────────────────────────

class Sexe(str, Enum):
    GARCON = "G"
    FILLE  = "F"

class Niveau(str, Enum):
    TRES_BON    = "TB"
    BON         = "B"
    MOYEN       = "M"
    FRAGILE     = "F"
    INCONNU     = "?"

class TypeAide(str, Enum):
    PAP  = "PAP"
    PAI  = "PAI"
    PPS  = "PPS"
    PPRE = "PPRE"
    AUCUNE = ""


# ─────────────────────────────────────────
#  Eleve
# ─────────────────────────────────────────

class Eleve(BaseModel):
    """
    Représentation complète d'un élève tel qu'extrait
    des tableaux de commission puis enrichi manuellement.
    """
    nom:           str
    prenom:        str
    sexe:          Sexe
    ecole:         str                   # nom de l'école d'origine
    commune:       Optional[str] = None  # commune de l'école

    niveau:        Niveau = Niveau.INCONNU
    aide:          TypeAide = TypeAide.AUCUNE

    est_moteur:    bool = False
    est_penible:   bool = False
    est_bilangue:  bool = False
    est_basket:    bool = False

    # Noms de famille uniquement (contrainte métier)
    a_separer_de:     list[str] = []
    a_regrouper_avec: list[str] = []

    observation:   str = ""   # texte brut corrigé, complété manuellement
    college:       str = "SMAC"          # collège d'affectation (SMAC par défaut)
    classe:        Optional[str] = None  # assignée lors de la répartition

    @field_validator("nom", "prenom", mode="before")
    @classmethod
    def strip_and_upper_nom(cls, v: str) -> str:
        return v.strip()

    @field_validator("a_separer_de", "a_regrouper_avec", mode="before")
    @classmethod
    def ensure_list(cls, v) -> list[str]:
        if isinstance(v, str):
            return [x.strip() for x in v.split(",") if x.strip()]
        return v or []

    @property
    def nom_complet(self) -> str:
        return f"{self.nom.upper()} {self.prenom.capitalize()}"

    @property
    def identifiant(self) -> str:
        """Clé de déduplication : NOM+PRENOM+ECOLE normalisés."""
        return f"{self.nom.upper()}_{self.prenom.upper()}_{self.ecole.upper()}"


# ─────────────────────────────────────────
#  EcoleCommission
# ─────────────────────────────────────────

class EcoleCommission(BaseModel):
    """
    Données brutes d'un tableau de commission
    reçu d'un enseignant référent.
    """
    nom_ecole:  str
    commune:    Optional[str] = None
    eleves:     list[Eleve] = []
    fichier_source: Optional[str] = None  # nom du fichier uploadé


# ─────────────────────────────────────────
#  Classe
# ─────────────────────────────────────────

class Classe(BaseModel):
    """
    Une classe de 6e, telle que configurée par le directeur
    et peuplée par l'algorithme de répartition.
    """
    nom:             str          # ex. "61", "62" …
    est_bilangue:    bool = False
    est_basket:      bool = False
    effectif_cible:  Optional[int] = None
    eleves:          list[Eleve] = []

    @property
    def effectif(self) -> int:
        return len(self.eleves)

    @property
    def nb_garcons(self) -> int:
        return sum(1 for e in self.eleves if e.sexe == Sexe.GARCON)

    @property
    def nb_filles(self) -> int:
        return sum(1 for e in self.eleves if e.sexe == Sexe.FILLE)

    @property
    def repartition_niveaux(self) -> dict[str, int]:
        counts: dict[str, int] = {n.value: 0 for n in Niveau}
        for e in self.eleves:
            counts[e.niveau.value] += 1
        return counts


# ─────────────────────────────────────────
#  ConfigRepartition
# ─────────────────────────────────────────

class ConfigRepartition(BaseModel):
    """
    Paramètres saisis par l'utilisateur dans la page Répartition.
    """
    classes:            list[Classe] = []
    noms_classes:       list[str] = ["61", "62", "63", "64", "65"]
    classes_bilangues:  list[str] = []
    classes_basket:     list[str] = []
    effectif_cible:     Optional[int] = None  # par défaut commun