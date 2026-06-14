"""
utils/toast.py
Messages toast standardisés pour toute l'application.
Chaque toast décrit précisément ce qui vient de se passer.
"""

import streamlit as st


def succes(message: str) -> None:
    st.toast(f"✅ {message}", icon="✅")


def erreur(message: str) -> None:
    st.toast(f"❌ {message}", icon="❌")


def info(message: str) -> None:
    st.toast(f"ℹ️ {message}", icon="ℹ️")


def avertissement(message: str) -> None:
    st.toast(f"⚠️ {message}", icon="⚠️")


# ─── Messages prédéfinis (cohérence terminologique) ───────────────────────────

def fichier_importe(nom: str, n_eleves: int) -> None:
    succes(f"{nom} importé — {n_eleves} élèves extraits.")


def fichier_deja_present(nom: str) -> None:
    avertissement(f"{nom} déjà importé, ignoré.")


def docx_genere(n_eleves: int) -> None:
    succes(f"synthese.docx généré — {n_eleves} élèves.")


def xlsx_genere(n_eleves: int, n_classes: int) -> None:
    succes(f"synthese.xlsx généré — {n_eleves} élèves, {n_classes} classes.")


def repartition_terminee(classes: list) -> None:
    detail = " · ".join(f"{c.nom} ({c.effectif})" for c in classes)
    succes(f"Répartition terminée : {detail}")


def contrainte_ignoree(nom_eleve: str, raison: str) -> None:
    avertissement(f"{nom_eleve} : contrainte ignorée ({raison}).")