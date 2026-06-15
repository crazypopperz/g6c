"""
app.py
Point d'entrée Streamlit — routing et accueil uniquement.
Toute logique métier est dans core/ et utils/.
"""
import streamlit as st
from utils.session import init_session, get

# ─── Configuration de la page ────────────────────────────────────────────────
st.set_page_config(
    page_title="G6C · Gestion 6e",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

init_session()

# ─── CSS global amélioré ────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ── En-tête principal ─────────────────────────────────────────────── */
    .main-header {
        background: linear-gradient(135deg, #0f0f0f 0%, #1a1a1a 100%);
        border-bottom: 3px solid #39FF14;
        padding: 2rem;
        border-radius: 0 0 15px 15px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(57, 255, 20, 0.15);
    }
    
    .main-header h1 {
        color: #39FF14;
        margin: 0;
        font-size: 2.2rem;
        font-weight: 800;
        letter-spacing: -1px;
    }
    
    .main-header p {
        color: #e0e0e0;
        margin: 0.5rem 0 0 0;
        font-size: 1.1rem;
    }

    /* ── Sidebar : indicateur d'étape active ──────────────────────────── */
    [data-testid="stSidebarNav"] a[aria-current="page"] {
        background: rgba(57, 255, 20, 0.12);
        border-left: 3px solid #39FF14;
        border-radius: 0 6px 6px 0;
        font-weight: 600;
    }

    /* ── Boutons primaires : neon green ───────────────────────────────── */
    .stButton > button[kind="primary"] {
        background: #39FF14;
        color: #0f0f0f;
        border: none;
        font-weight: 700;
        letter-spacing: 0.03em;
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
        transition: all 0.3s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background: #2ecc10;
        color: #0f0f0f;
        box-shadow: 0 4px 12px rgba(57, 255, 20, 0.4);
        transform: translateY(-2px);
    }

    /* ─ Tableau des étapes ───────────────────────────────────────────── */
    .step-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        margin: 1rem 0;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    .step-table th {
        background: #1a1a1a;
        color: #39FF14;
        padding: 1rem;
        text-align: left;
        font-weight: 700;
    }
    .step-table td {
        padding: 1rem;
        border-bottom: 1px solid #f0f0f0;
        background: white;
    }
    .step-table tr:last-child td {
        border-bottom: none;
    }
    .step-table tr:hover td {
        background: #f9fff9;
    }

    /* ─ Badges de niveau ─────────────────────────────────────────────── */
    .badge-TB { background:#1a5c1a; color:#9effa0; padding:2px 8px; border-radius:4px; font-weight:700; }
    .badge-B  { background:#2d6b2d; color:#b8ffba; padding:2px 8px; border-radius:4px; font-weight:700; }
    .badge-M  { background:#6b5a00; color:#ffe680; padding:2px 8px; border-radius:4px; font-weight:700; }
    .badge-F  { background:#6b1a1a; color:#ffaaaa; padding:2px 8px; border-radius:4px; font-weight:700; }

    /* ── Mention bilangue ─────────────────────────────────────────────── */
    .bilangue { color:#39FF14; font-weight:700; font-size:0.85em; margin-left:6px; }

    /* ── Metrics cards ────────────────────────────────────────────────── */
    [data-testid="stMetric"] {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        border: 1px solid #f0f0f0;
    }
    [data-testid="stMetricLabel"] {
        color: #666 !important;
        font-weight: 500 !important;
        font-size: 0.9rem !important;
    }
    [data-testid="stMetricValue"] {
        color: #0f0f0f !important;
        font-weight: 800 !important;
        font-size: 1.8rem !important;
    }
    [data-testid="stMetricDelta"] {
        color: #39FF14 !important;
    }
</style>
""", unsafe_allow_html=True)

# ─── En-tête visuel ──────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🏫 G6C — Gestion des commissions CM2</h1>
    <p>Automatisez l'extraction, la synthèse et la répartition de vos élèves de 6e.</p>
</div>
""", unsafe_allow_html=True)

# ─── Instructions (Tableau stylé) ────────────────────────────────────────────
st.markdown("### ️ Feuille de route")
st.markdown("""
<div style="background: #f9fff9; padding: 1rem; border-radius: 10px; border-left: 4px solid #39FF14; margin-bottom: 2rem;">
    <p style="margin: 0; color: #333;">
        💡 <strong>Conseil :</strong> Vous pouvez entrer dans le processus à n'importe quelle étape. 
        Si vous avez déjà un <code>synthese.docx</code> modifié manuellement, allez directement en <strong>Étape 3</strong>.
    </p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<table class="step-table">
    <thead>
        <tr>
            <th>Étape</th>
            <th>Page</th>
            <th>Rôle</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td><strong>1</strong></td>
            <td>📥 Import</td>
            <td>Déposer les tableaux de commission (xlsx / xls / pdf / ods)</td>
        </tr>
        <tr>
            <td><strong>2</strong></td>
            <td>📝 Synthèse DOCX</td>
            <td>Générer <code>synthese.docx</code> avec code couleur et tableau par élève</td>
        </tr>
        <tr>
            <td><strong>3</strong></td>
            <td> Synthèse XLSX</td>
            <td>Fusionner les docx et produire <code>synthese.xlsx</code></td>
        </tr>
        <tr>
            <td><strong>4</strong></td>
            <td>🏫 Répartition</td>
            <td>Configurer les classes et lancer l'algorithme de répartition</td>
        </tr>
    </tbody>
</table>
""", unsafe_allow_html=True)

# ─── Résumé de session (Metrics) ─────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📈 État de la session en cours")

commissions = get("commissions")
eleves = get("eleves_fusionnes")
classes = get("classes_finales")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Tableaux importés", len(commissions) if commissions else 0)
with col2:
    st.metric("Élèves extraits", sum(len(c.eleves) for c in commissions) if commissions else 0)
with col3:
    st.metric("Élèves fusionnés", len(eleves) if eleves else 0)
with col4:
    st.metric("Classes constituées", len(classes) if classes else 0)

# ── Bouton Reset ─────────────────────────────────────────────────────────────
st.divider()
col_reset, col_info = st.columns([1, 3])

with col_reset:
    if st.button("🗑️ Reset", type="secondary", use_container_width=True):
        from utils.session import reset_all
        reset_all()
        st.rerun()

with col_info:
    st.caption("Videz toutes les données pour recommencer depuis zéro.")