"""
utils/session.py
Helpers pour gérer le session_state Streamlit de façon centralisée.
Toutes les clés utilisées dans l'application sont déclarées ici.
"""

import streamlit as st
from typing import Any


# ─────────────────────────────────────────
#  Clés de session (source of truth)
# ─────────────────────────────────────────

KEYS_DEFAULTS: dict[str, Any] = {
    # Sprint 1 — Import
    "commissions":          [],   # list[EcoleCommission] — données extraites
    "fichiers_importes":    [],   # list[str] — noms des fichiers déjà traités

    # Sprint 2 — Synthèse DOCX
    "docx_bytes":           None, # bytes du synthese.docx généré

    # Sprint 3 — Synthèse XLSX
    "eleves_fusionnes":     [],   # list[Eleve] — après fusion des docx
    "xlsx_bytes":           None, # bytes du synthese.xlsx

    # Sprint 4 — Répartition
    "config_repartition":   None, # ConfigRepartition
    "classes_finales":      [],   # list[Classe] après algo

    # Navigation
    "etape_courante":       1,    # 1-4, suit le flux principal
}


def init_session() -> None:
    """À appeler en tête de chaque page pour garantir l'initialisation."""
    for key, default in KEYS_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = default


def get(key: str) -> Any:
    init_session()
    return st.session_state[key]


def set(key: str, value: Any) -> None:
    st.session_state[key] = value


def reset_from(etape: int) -> None:
    """
    Réinitialise toutes les données à partir d'une étape donnée.
    Utile quand un nouvel import invalide les étapes suivantes.
    """
    if etape <= 2:
        st.session_state["docx_bytes"] = None
    if etape <= 3:
        st.session_state["eleves_fusionnes"] = []
        st.session_state["xlsx_bytes"] = None
    if etape <= 4:
        st.session_state["config_repartition"] = None
        st.session_state["classes_finales"] = []

def reset_all() -> None:
    """Réinitialise TOUTES les données de session."""
    for key, default in KEYS_DEFAULTS.items():
        st.session_state[key] = default
    st.toast("🔄 Session réinitialisée", icon="🔄")