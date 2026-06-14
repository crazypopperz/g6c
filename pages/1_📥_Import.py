"""
pages/1_📥_Import.py
Deux colonnes :
  - Gauche : premier import de tableaux de commission
  - Droite  : mise à jour d'un docx existant avec de nouveaux tableaux
"""

import streamlit as st
from pathlib import Path
import tempfile, os
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent.parent / ".env")

from utils.session import init_session, get, set as sset, reset_from
import utils.toast as toast
from core.file_parser import parse_file
from core.extractor import extract_from_text
from core.models import EcoleCommission

st.set_page_config(page_title="Import · G6C", page_icon="📥", layout="wide")
init_session()

st.title("📥 Étape 1 — Import des tableaux de commission")

# ══════════════════════════════════════════════════════════════════════════════
col_gauche, sep, col_droite = st.columns([10, 1, 10])

# ─── COLONNE GAUCHE : Premier import ─────────────────────────────────────────
with col_gauche:
    st.subheader("🆕 Première génération")
    st.caption(
        "Déposez ici les tableaux de commission reçus des enseignants. "
        "Vous pouvez en ajouter d'autres plus tard sans tout reprendre — "
        "les nouveaux seront simplement ajoutés à ceux déjà importés."
    )

    uploaded = st.file_uploader(
        "Tableaux de commission (xlsx, xls, ods, pdf)",
        type=["xlsx", "xls", "ods", "pdf"],
        accept_multiple_files=True,
        key="upload_tableaux",
        label_visibility="collapsed",
    )

    if uploaded:
        commissions: list = get("commissions")
        deja_importes: list = get("fichiers_importes")
        nouveaux = [f for f in uploaded if f.name not in deja_importes]

        if not nouveaux:
            st.info("Ces fichiers ont déjà été importés.")
        else:
            mistral_key = os.environ.get("MISTRAL_API_KEY", "")
            if not mistral_key:
                st.error("⚠️ Clé MISTRAL_API_KEY manquante.")
                st.stop()

            progress = st.progress(0, text="Préparation…")
            erreurs = []

            for i, fichier in enumerate(nouveaux):
                progress.progress(i / len(nouveaux), text=f"Lecture de {fichier.name}…")
                suffix = Path(fichier.name).suffix
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(fichier.read())
                    tmp_path = tmp.name
                try:
                    texte = parse_file(tmp_path)
                    progress.progress((i + 0.5) / len(nouveaux), text=f"Extraction : {fichier.name}…")
                    eleves = extract_from_text(texte, source_fichier=fichier.name)
                    meta_lines = {
                        line.split(":")[0].strip(): ":".join(line.split(":")[1:]).strip()
                        for line in texte.splitlines()
                        if ":" in line and not line.startswith("ELEVE")
                    }
                    commission = EcoleCommission(
                        nom_ecole      = meta_lines.get("ECOLE", fichier.name),
                        commune        = meta_lines.get("COMMUNE", ""),
                        eleves         = eleves,
                        fichier_source = fichier.name,
                    )
                    commissions.append(commission)
                    deja_importes.append(fichier.name)
                    toast.fichier_importe(fichier.name, len(eleves))
                except Exception as exc:
                    erreurs.append(f"{fichier.name} : {exc}")
                    toast.erreur(f"Échec sur {fichier.name}")
                finally:
                    os.unlink(tmp_path)

            progress.progress(1.0, text="Import terminé.")
            sset("commissions", commissions)
            sset("fichiers_importes", deja_importes)
            reset_from(2)

            if erreurs:
                st.warning("Fichiers non traités :\n" + "\n".join(erreurs))

    # Récapitulatif
    commissions = get("commissions")
    if commissions:
        st.divider()
        st.subheader(f"{len(commissions)} tableau(x) importé(s)")
        for comm in commissions:
            with st.expander(f"**{comm.nom_ecole}** — {comm.commune} ({len(comm.eleves)} élèves)"):
                rows = []
                for e in comm.eleves:
                    rows.append({
                        "Nom":        e.nom,
                        "Prénom":     e.prenom,
                        "Sexe":       e.sexe.value,
                        "Niveau":     e.niveau.value,
                        "Aide":       e.aide.value,
                        "Bilangue":   "✓" if e.est_bilangue else "",
                        "Basket":     "✓" if e.est_basket   else "",
                        "Pénible":    "✓" if e.est_penible  else "",
                        "Collège":    e.college,
                        "Observation": e.observation[:80] + "…" if len(e.observation) > 80 else e.observation,
                    })
                st.dataframe(rows, width="stretch", hide_index=True)

        total = sum(len(c.eleves) for c in commissions)
        st.metric("Total élèves extraits", total)
        st.info("👈 Rendez-vous sur **📝 Synthese DOCX** dans la barre latérale pour générer votre document.")
    else:
        st.info("Aucun tableau importé. Déposez vos fichiers ci-dessus.")

# ─── SÉPARATEUR VISUEL ────────────────────────────────────────────────────────
with sep:
    st.markdown(
        "<div style='border-left:2px solid #444; height:100%; margin:0 auto;'></div>",
        unsafe_allow_html=True
    )

# ─── COLONNE DROITE : Mise à jour ────────────────────────────────────────────
with col_droite:
    st.subheader("🔄 Mise à jour d'un docx existant")
    st.caption(
        "Vous avez déjà généré et annoté un `synthese.docx` ? "
        "Déposez-le ici avec les nouveaux tableaux de commission reçus. "
        "L'application fusionnera tout sans doublon et vous rendra "
        "un nouveau docx complet avec vos annotations préservées."
    )

    uploaded_docx = st.file_uploader(
        "Votre synthese.docx annoté",
        type=["docx"],
        key="upload_docx_existant",
        label_visibility="visible",
    )

    uploaded_nouveaux = st.file_uploader(
        "Nouveaux tableaux de commission",
        type=["xlsx", "xls", "ods", "pdf"],
        accept_multiple_files=True,
        key="upload_nouveaux_tableaux",
        label_visibility="visible",
    )

    if uploaded_docx and uploaded_nouveaux:
        if st.button("🔄 Fusionner et générer le nouveau docx", type="primary"):
            mistral_key = os.environ.get("MISTRAL_API_KEY", "")
            if not mistral_key:
                st.error("⚠️ Clé MISTRAL_API_KEY manquante.")
                st.stop()

            with st.spinner("Extraction du docx existant…"):
                from core.docx_extractor import extract_eleves_from_docx, fusionner_eleves
                from core.docx_builder import build_synthese_docx
                from core.models import EcoleCommission

                # 1. Extraire les élèves du docx existant
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp.write(uploaded_docx.read())
                    tmp_docx = tmp.name
                try:
                    eleves_anciens = extract_eleves_from_docx(tmp_docx)
                finally:
                    os.unlink(tmp_docx)

            eleves_nouveaux = []
            erreurs = []
            progress2 = st.progress(0, text="Préparation des nouveaux tableaux…")
            for i, fichier in enumerate(uploaded_nouveaux):
                progress2.progress(i / len(uploaded_nouveaux), text=f"Extraction : {fichier.name}…")
                suffix = Path(fichier.name).suffix
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(fichier.read())
                    tmp_path = tmp.name
                try:
                    texte  = parse_file(tmp_path)
                    eleves = extract_from_text(texte, source_fichier=fichier.name)
                    eleves_nouveaux.extend(eleves)
                    toast.fichier_importe(fichier.name, len(eleves))
                except Exception as exc:
                    erreurs.append(f"{fichier.name} : {exc}")
                finally:
                    os.unlink(tmp_path)
            progress2.progress(1.0, text="Extraction terminée.")

            with st.spinner("Fusion et génération du nouveau docx…"):
                eleves_fusionnes = fusionner_eleves(eleves_anciens, eleves_nouveaux)

                # Regrouper par école pour build_synthese_docx
                ecoles: dict = {}
                for e in eleves_fusionnes:
                    key = e.ecole or "École inconnue"
                    if key not in ecoles:
                        ecoles[key] = EcoleCommission(
                            nom_ecole=key,
                            commune=e.commune or "",
                            eleves=[]
                        )
                    ecoles[key].eleves.append(e)

                docx_bytes = build_synthese_docx(list(ecoles.values()))

            st.success(f"✅ Fusion réussie — {len(eleves_fusionnes)} élèves ({len(eleves_anciens)} anciens + {len(eleves_nouveaux)} nouveaux)")

            if erreurs:
                st.warning("Fichiers non traités :\n" + "\n".join(erreurs))

            st.download_button(
                label="⬇️ Télécharger le nouveau synthese.docx",
                data=docx_bytes,
                file_name="synthese.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )

    elif uploaded_docx and not uploaded_nouveaux:
        st.info("Déposez maintenant les nouveaux tableaux de commission.")
    elif uploaded_nouveaux and not uploaded_docx:
        st.info("Déposez votre synthese.docx annoté pour lancer la fusion.")
    else:
        st.info("Déposez votre docx et les nouveaux tableaux pour lancer la fusion.")