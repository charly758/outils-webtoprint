import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
import io

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Portail Web-to-Print", page_icon="🚀", layout="centered")

# --- CSS PERSONNALISÉ (Pour un look plus propre et pro) ---
st.markdown("""
    <style>
        /* Masquer le menu technique Streamlit en haut à droite et le pied de page */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* Styliser un peu le titre principal */
        .main-title {
            font-size: 2.5rem;
            font-weight: 700;
            color: #1E3A8A;
            margin-bottom: 0px;
        }
        .subtitle {
            font-size: 1.1rem;
            color: #6B7280;
            margin-bottom: 30px;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. FONCTION : FACTURATION
# ==========================================
def trouver_categorie(row):
    produit = str(row.get('Product Name', '')).lower()
    pays = str(row.get('Ship Country Name', '')).strip().lower()
    
    if pays and pays != 'france' and pays != 'nan' and pays != '':
        if pays == 'canada': return "NAM"
        return "Export"
    
    promos_absolues = ["coffret rituals", "mug", "roulette à pizza", "cubicool", "catalogue bizline france"]
    if any(p in produit for p in promos_absolues): return "Promo"
    
    BIZSHOW = ["biz'show", "carte cadeau total 25€"]
    MERCH = ["biz 99", "présentoir", "tablette", "cartobacs", "fronton", "kit fond", "bac", "mannequin", "support pour rubans", "plaque perforée"]
    GOODIES = ["sweat", "t-shirt", "polo", "casquette", "tote bag", "stylo", "fouta", "biscuits", "jeu de cartes", "sachet de bonbon", "echarpe"]
    PROMO = ["instax", "rose éternelle", "roundnet", "bracelet"]
    ZEBRA = ["zebra", "m&m's"]
    RH = ["cartes de visite"]
    
    if any(m in produit for m in BIZSHOW): return "Biz'Show"
    if any(m in produit for m in MERCH): return "Merchandising"
    if any(m in produit for m in GOODIES): return "Goodies"
    if any(m in produit for m in PROMO): return "Promo"
    if any(m in produit for m in ZEBRA): return "Zebra"
    if any(m in produit for m in RH): return "RH"
    return ""

def traiter_facturation(file):
    df = pd.read_excel(file)
    df.columns = df
