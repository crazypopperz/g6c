"""
pages/3_📊_Synthese_XLSX.py
Génère synthese.xlsx depuis synthese.docx.
Extraction Python pure via xlsx_builder (sans Mistral).
"""
import streamlit as st
from pathlib import Path
import tempfile
from core.xlsx_builder import build_synthese_xlsx

st.set_page_config(page_title="Synthèse XLSX", page_icon="📊", layout="wide")

st.title("📊 Étape 3 — Génération de la synthèse Excel")

st.markdown("""
Cette page convertit le fichier **synthese.docx** (généré à l'étape 2) 
en un fichier **synthese.xlsx** prêt pour la répartition.

Le fichier Excel contient :
- Nom, Prénom, Sexe, Ville
- Bilangue, Basket, Moteur, Pénible
- Aide (PPRE/PAP/PAI/PPS)
- À séparer de / À regrouper avec
""")

st.divider()

# ── Upload du DOCX ────────────────────────────────────────────────────────────
docx_file = st.file_uploader(
    "📄 Upload le fichier synthese.docx",
    type=["docx"],
    help="Fichier généré à l'étape 2"
)

if docx_file is None:
    st.info("👆 Commence par uploader ton fichier synthese.docx")
    st.stop()

# ── Aperçu ───────────────────────────────────────────────────────────────────
st.success(f"✅ Fichier chargé : **{docx_file.name}** ({docx_file.size / 1024:.1f} Ko)")

if st.button("⚙️ Générer le fichier Excel", type="primary", use_container_width=True):
    with st.spinner("Génération en cours..."):
        try:
            # Sauvegarde temporaire du DOCX
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(docx_file.read())
                tmp_docx_path = Path(tmp_docx.name)
            
            # Chemin de sortie
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
                tmp_xlsx_path = Path(tmp_xlsx.name)
            
            # Génération
            build_synthese_xlsx(tmp_docx_path, tmp_xlsx_path)
            
            # Lecture du résultat
            with open(tmp_xlsx_path, "rb") as f:
                xlsx_bytes = f.read()
            
            # Nettoyage
            tmp_docx_path.unlink()
            tmp_xlsx_path.unlink()
            
            # Stockage en session pour le download
            st.session_state["xlsx_bytes"] = xlsx_bytes
            st.session_state["xlsx_ready"] = True
            
            st.success("✅ Fichier Excel généré avec succès !")
            
        except Exception as e:
            st.error(f"❌ Erreur lors de la génération : {e}")
            st.exception(e)

# ── Bouton de téléchargement ──────────────────────────────────────────────────
if st.session_state.get("xlsx_ready"):
    st.divider()
    st.subheader("⬇️ Télécharger le fichier")
    
    st.download_button(
        label="📥 Télécharger synthese.xlsx",
        data=st.session_state["xlsx_bytes"],
        file_name="synthese.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    
    st.info("💡 Prochaine étape : va à la page **4_🏫_Répartition** pour créer les classes.")