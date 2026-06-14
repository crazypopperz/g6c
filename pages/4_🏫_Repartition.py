"""
pages/4_🏫_Repartition.py
Configuration des classes et répartition des élèves.
Interface professionnelle et épurée.
"""
import streamlit as st
from pathlib import Path
import tempfile
import pandas as pd
from datetime import datetime
from core.distributor import (
    lire_eleves_depuis_xlsx,
    repartir_eleves,
    generer_xlsx_repartition,
    Classe,
)
from core.styles import inject_css, page_header

st.set_page_config(page_title="Répartition", page_icon="🏫", layout="wide")
inject_css()

# ── Initialisation session_state ──────────────────────────────────────────────
if "classes_noms" not in st.session_state:
    st.session_state["classes_noms"] = []
if "classes_basket" not in st.session_state:
    st.session_state["classes_basket"] = []
if "classes_bilangue" not in st.session_state:
    st.session_state["classes_bilangue"] = []
if "effectif_cible" not in st.session_state:
    st.session_state["effectif_cible"] = 28
if "repartition_done" not in st.session_state:
    st.session_state["repartition_done"] = False
if "xlsx_bytes" not in st.session_state:
    st.session_state["xlsx_bytes"] = None
if "classes_result" not in st.session_state:
    st.session_state["classes_result"] = []

# ── En-tête ───────────────────────────────────────────────────────────────────
page_header("🏫 Répartition des classes", "Configurez vos classes et lancez l'algorithme de répartition")

# ── Upload du synthese.xlsx ───────────────────────────────────────────────────
col_upload, col_info = st.columns([3, 1])
with col_upload:
    xlsx_file = st.file_uploader(
        "Fichier source synthese.xlsx",
        type=["xlsx"],
        label_visibility="collapsed",
        help="Fichier généré à l'étape 3",
    )

if xlsx_file is None:
    st.info("⬆️ Commencez par uploader votre fichier **synthese.xlsx** ci-dessus.")
    st.stop()

with col_info:
    st.success(f"✅ {xlsx_file.name}")

# ── Deux colonnes : Configuration | Validation ───────────────────────────────
col_config, col_validation = st.columns([5, 7], gap="large")

# ══════════════════════════════════════════════════════════════════════════════
# COLONNE GAUCHE : CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════
with col_config:
    st.subheader("⚙️ Paramètres de répartition")

    # 1. Noms des classes
    with st.container(border=True):
        st.markdown("**1. Création des classes**")
        saisie_classes = st.text_input(
            "Noms des classes (séparés par des virgules)",
            value=", ".join(st.session_state["classes_noms"]) if st.session_state["classes_noms"] else "",
            placeholder="Ex: 6A, 6B, 6C, 6D",
            label_visibility="collapsed",
        )

        if st.button("Valider les classes", type="primary", use_container_width=True):
            noms = [n.strip() for n in saisie_classes.split(",") if n.strip()]
            if noms:
                st.session_state["classes_noms"] = noms
                st.session_state["classes_basket"] = []
                st.session_state["classes_bilangue"] = []
                st.session_state["repartition_done"] = False
                st.rerun()
            else:
                st.error("Veuillez entrer au moins un nom de classe.")

    # 2. Classes Basket & Bilangue
    with st.container(border=True):
        st.markdown("**2. Classes spécialisées**")
        
        if st.session_state["classes_noms"]:
            c1, c2 = st.columns(2)
            with c1:
                basket_selected = st.multiselect(
                    "Section Basket 🏀",
                    options=st.session_state["classes_noms"],
                    default=st.session_state["classes_basket"],
                    key="basket_select",  # ← AJOUTER UNE KEY UNIQUE
                    label_visibility="visible",
                )
                st.session_state["classes_basket"] = basket_selected

            with c2:
                bilangue_selected = st.multiselect(
                    "Classes Bilangues 🇬🇧",
                    options=st.session_state["classes_noms"],
                    default=st.session_state["classes_bilangue"],
                    key="bilangue_select",  # ← AJOUTER UNE KEY UNIQUE
                    label_visibility="visible",
                )
                st.session_state["classes_bilangue"] = bilangue_selected
        else:
            st.warning("Validez d'abord les noms de classes.")

    # 3. Effectif cible
    with st.container(border=True):
        st.markdown("**3. Effectif cible**")
        effectif = st.number_input(
            "Élèves par classe",
            min_value=15, max_value=40,
            value=st.session_state["effectif_cible"],
            step=1,
            label_visibility="collapsed",
        )
        st.session_state["effectif_cible"] = effectif

    st.divider()

    # Bouton de lancement
    if st.button("🚀 Lancer la répartition", type="primary", use_container_width=True, icon="🔀"):
        if not st.session_state["classes_noms"]:
            st.error("Veuillez d'abord valider les noms de classes.")
        else:
            with st.spinner("Algorithme de répartition en cours..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        tmp.write(xlsx_file.read())
                        tmp_path = Path(tmp.name)

                    eleves = lire_eleves_depuis_xlsx(tmp_path)
                    
                    classes = [
                        Classe(
                            nom=nom,
                            est_basket=nom in st.session_state["classes_basket"],
                            est_bilangue=nom in st.session_state["classes_bilangue"],
                        )
                        for nom in st.session_state["classes_noms"]
                    ]

                    classes = repartir_eleves(eleves, classes, st.session_state["effectif_cible"])

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
                        tmp_out_path = Path(tmp_out.name)

                    annee = datetime.now().year
                    generer_xlsx_repartition(classes, tmp_out_path, str(annee))

                    with open(tmp_out_path, "rb") as f:
                        st.session_state["xlsx_bytes"] = f.read()

                    st.session_state["repartition_done"] = True
                    st.session_state["classes_result"] = classes

                    tmp_path.unlink()
                    tmp_out_path.unlink()
                    st.success("✅ Répartition terminée !")

                except Exception as e:
                    st.error(f"❌ Erreur : {e}")

# ═════════════════════════════════════════════════════════════════════════════
# COLONNE DROITE : VALIDATION & RÉSULTATS
# ═════════════════════════════════════════════════════════════════════════════
with col_validation:
    st.subheader("📋 Récapitulatif de la configuration")

    if not st.session_state["classes_noms"]:
        st.empty()
    else:
        # Tableau récapitulatif pro
        recap_data = []
        for nom in st.session_state["classes_noms"]:
            badges = []
            if nom in st.session_state["classes_basket"]:
                badges.append("Basket")
            if nom in st.session_state["classes_bilangue"]:
                badges.append("Bilangue")
            
            recap_data.append({
                "Classe": nom,
                "Spécialités": ", ".join(badges) if badges else "Standard",
            })
        
        df_recap = pd.DataFrame(recap_data)
        st.dataframe(df_recap, hide_index=True, use_container_width=True, column_config={
            "Classe": st.column_config.TextColumn("Nom de la classe"),
            "Spécialités": st.column_config.TextColumn("Type de classe"),
        })

        st.metric("Effectif cible", f"{st.session_state['effectif_cible']} élèves/classe")

    # Résultat de la répartition
    if st.session_state.get("repartition_done") and st.session_state.get("classes_result"):
        st.divider()
        st.subheader("📊 Résultats de la répartition")

        resultats = []
        for c in st.session_state["classes_result"]:
            compteurs = c.compteurs_niveaux()
            resultats.append({
                "Classe": c.nom,
                "Effectif": c.effectif,
                "Garçons": c.nb_garcons,
                "Filles": c.nb_filles,
                "TB / B / M / F": f"{compteurs['TB']} / {compteurs['B']} / {compteurs['M']} / {compteurs['F']}",
                "Spécialité": ("Basket" if c.est_basket else "") + (" + Bilangue" if c.est_bilangue else ""),
            })
        
        df_res = pd.DataFrame(resultats)
        st.dataframe(df_res, hide_index=True, use_container_width=True, column_config={
            "Classe": st.column_config.TextColumn("Classe"),
            "Effectif": st.column_config.NumberColumn("Total"),
            "Garçons": st.column_config.NumberColumn("G"),
            "Filles": st.column_config.NumberColumn("F"),
        })

    # Téléchargement
    if st.session_state.get("xlsx_bytes"):
        st.divider()
        annee = datetime.now().year
        st.download_button(
            label=f"📥 Télécharger repartition_{annee}.xlsx",
            data=st.session_state["xlsx_bytes"],
            file_name=f"repartition_{annee}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            icon="💾",
        )