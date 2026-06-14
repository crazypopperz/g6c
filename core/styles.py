"""
core/styles.py
Styles CSS réutilisables pour toutes les pages de l'application.
"""

# CSS global à injecter dans chaque page avec st.markdown(CSS_GLOBAL, unsafe_allow_html=True)
CSS_GLOBAL = """
<style>
    /* ── Titres de page ────────────────────────────────────────────────── */
    .page-header {
        background: linear-gradient(135deg, #1A5C1A 0%, #4CAF50 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 15px rgba(26, 92, 26, 0.2);
    }
    
    .page-header h1 {
        margin: 0;
        font-size: 1.8rem;
        font-weight: 700;
    }
    
    .page-header p {
        margin: 0.3rem 0 0 0;
        opacity: 0.95;
    }
    
    /* ── Conteneurs bordurés ───────────────────────────────────────────── */
    .stContainer {
        border-radius: 10px !important;
    }
    
    /* ── Boutons ───────────────────────────────────────────────────────── */
    .stButton > button {
        background-color: #1A5C1A !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.5rem 1.5rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        background-color: #2E7D32 !important;
        box-shadow: 0 4px 12px rgba(26, 92, 26, 0.3) !important;
        transform: translateY(-1px) !important;
    }
    
    /* ── File uploader ─────────────────────────────────────────────────── */
    .stFileUploader {
        border: 2px dashed #1A5C1A !important;
        border-radius: 10px !important;
    }
    
    /* ── Metrics ───────────────────────────────────────────────────────── */
    [data-testid="stMetric"] {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    }
    
    [data-testid="stMetricLabel"] {
        color: #6B7280 !important;
        font-weight: 500 !important;
    }
    
    [data-testid="stMetricValue"] {
        color: #1A5C1A !important;
        font-weight: 700 !important;
    }
    
    /* ── Sidebar ───────────────────────────────────────────────────────── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #F5F7F5 0%, #FFFFFF 100%);
    }
    
    /* ── Masquer le menu Streamlit ─────────────────────────────────────── */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
"""

def inject_styles():
    """Injecte le CSS global dans la page courante."""
    import streamlit as st
    st.markdown(CSS_GLOBAL, unsafe_allow_html=True)

def page_header(title: str, subtitle: str = "", icon: str = ""):
    """Affiche un en-tête de page stylé."""
    import streamlit as st
    
    display_title = f"{icon} {title}" if icon else title
    
    st.markdown(f"""
    <div class="page-header">
        <h1>{display_title}</h1>
        {"<p>" + subtitle + "</p>" if subtitle else ""}
    </div>
    """, unsafe_allow_html=True)
# Alias pour compatibilité
inject_css = inject_styles