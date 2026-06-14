"""
pages/2_📝_Synthese_DOCX.py
Génération de synthese.docx depuis les commissions importées en session.
"""

import streamlit as st
from dotenv import load_dotenv
from pathlib import Path
load_dotenv(Path(__file__).parent.parent / ".env")

from utils.session import init_session, get, set as sset
import utils.toast as toast
from core.docx_builder import build_synthese_docx
from core.styles import inject_css, page_header

st.set_page_config(page_title="Synthèse DOCX · G6C", page_icon="📝", layout="wide")
init_session()
inject_css()

page_header("📝 Synthèse DOCX", "Génération du document Word récapitulatif")

commissions = get("commissions")

if not commissions:
    st.warning("Aucun tableau importé. Retournez à l'étape 1.")
    st.stop()

# ── Résumé ────────────────────────────────────────────────────────────────────
total = sum(len(c.eleves) for c in commissions)
col1, col2 = st.columns(2)
col1.metric("Écoles importées", len(commissions))
col2.metric("Élèves au total", total)

st.divider()

# ── Génération ────────────────────────────────────────────────────────────────
docx_bytes = get("docx_bytes")

if st.button("⚙️ Générer synthese.docx", type="primary"):
    with st.spinner("Génération en cours…"):
        try:
            docx_bytes = build_synthese_docx(commissions)
            sset("docx_bytes", docx_bytes)
            toast.docx_genere(total)
        except Exception as e:
            st.error(f"Erreur lors de la génération : {e}")
            st.stop()

if docx_bytes:
    st.success(f"✅ synthese.docx prêt — {total} élèves")
    st.download_button(
        label="⬇️ Télécharger synthese.docx",
        data=docx_bytes,
        file_name="synthese.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
    st.divider()
    st.info(
        "**Étape suivante :**\n\n"
        "1. Ouvrez `synthese.docx` dans Word\n"
        "2. Complétez ou corrigez les observations si besoin\n"
        "3. Enregistrez le fichier\n"
        "4. Passez à l'**Étape 3** pour générer le fichier Excel final"
    )
else:
    st.info("Cliquez sur le bouton ci-dessus pour générer le document.")