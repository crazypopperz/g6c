"""
app.py
Point d'entrée Streamlit — routing et accueil uniquement.
Toute logique métier est dans core/ et utils/.
"""

import streamlit as st
from utils.session import init_session

# ─── Configuration de la page ─────────────────────────────────────────────────
st.set_page_config(
    page_title    = "G6C · Gestion 6e",
    page_icon     = "🏫",
    layout        = "wide",
    initial_sidebar_state = "expanded",
)

init_session()

# ─── CSS global ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Sidebar : indicateur d'étape active */
[data-testid="stSidebarNav"] a[aria-current="page"] {
    background: rgba(57, 255, 20, 0.12);
    border-left: 3px solid #39FF14;
    border-radius: 0 6px 6px 0;
    font-weight: 600;
}

/* Boutons primaires : neon green */
.stButton > button[kind="primary"] {
    background: #39FF14;
    color: #0f0f0f;
    border: none;
    font-weight: 700;
    letter-spacing: 0.03em;
}
.stButton > button[kind="primary"]:hover {
    background: #2ecc10;
    color: #0f0f0f;
}

/* Badges de niveau dans les tableaux */
.badge-TB { background:#1a5c1a; color:#9effa0; padding:2px 8px; border-radius:4px; font-weight:700; }
.badge-B  { background:#2d6b2d; color:#b8ffba; padding:2px 8px; border-radius:4px; font-weight:700; }
.badge-M  { background:#6b5a00; color:#ffe680; padding:2px 8px; border-radius:4px; font-weight:700; }
.badge-F  { background:#6b1a1a; color:#ffaaaa; padding:2px 8px; border-radius:4px; font-weight:700; }

/* Mention bilangue */
.bilangue { color:#39FF14; font-weight:700; font-size:0.85em; margin-left:6px; }
</style>
""", unsafe_allow_html=True)

# ─── Accueil ──────────────────────────────────────────────────────────────────
st.title("🏫 G6C — Gestion des commissions CM2")
st.markdown("""
Bienvenue dans G6C. Suivez les étapes dans l'ordre ou accédez directement
à celle dont vous avez besoin.

| Étape | Page | Rôle |
|-------|------|------|
| 1 | 📥 Import | Déposer les tableaux de commission (xlsx / xls / pdf / ods) |
| 2 | 📝 Synthèse DOCX | Générer `synthese.docx` avec code couleur et tableau par élève |
| 3 | 📊 Synthèse XLSX | Fusionner les docx et produire `synthese.xlsx` |
| 4 | 🏫 Répartition | Configurer les classes et lancer la répartition |

> **Conseil :** vous pouvez entrer dans le processus à n'importe quelle étape.  
> Si vous avez déjà un `synthese.docx` modifié, allez directement en **Étape 3**.
""")

# ─── Résumé de session ────────────────────────────────────────────────────────
from utils.session import get

commissions = get("commissions")
eleves = get("eleves_fusionnes")
classes = get("classes_finales")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Tableaux importés",  len(commissions))
col2.metric("Élèves extraits",    sum(len(c.eleves) for c in commissions) if commissions else 0)
col3.metric("Élèves fusionnés",   len(eleves))
col4.metric("Classes constituées", len(classes))