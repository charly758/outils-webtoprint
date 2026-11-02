import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
import io
import os

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Portail Web-to-Print", page_icon="🚀", layout="centered")

# --- CSS PERSONNALISÉ ---
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .main-title { 
            font-size: 2.5rem; 
            font-weight: 700; 
            color: #222222; /* Noir Microlynx */
            margin-bottom: 0px; 
        }
        .subtitle { 
            font-size: 1.1rem; 
            color: #BA2025; /* Rouge Agent Paper */
            margin-bottom: 30px; 
            font-weight: 600;
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
    df.columns = df.columns.str.strip() 
    colonnes_requises = ['Customer Name', 'Order ID', 'Product Name', 'Total Quantity', 'DateOrderCreated', 'Ship_Company', 'Ship_City', 'Ship Country Name', 'Total Order Price']
    
    if 'Product Name' not in df.columns: raise ValueError("Colonne 'Product Name' introuvable.")

    df['Catégorie'] = df.apply(trouver_categorie, axis=1)
    df = df[[c for c in colonnes_requises + ['Catégorie'] if c in df.columns]].copy()
    
    df['DateOrderCreated'] = pd.to_datetime(df['DateOrderCreated'], errors='coerce').dt.strftime('%d/%m/%Y')
    df.sort_values(by=['DateOrderCreated', 'Order ID'], inplace=True)
    
    mask_etranger = df['Ship Country Name'].astype(str).str.strip().str.lower().apply(lambda x: x not in ['france', 'nan', ''])
    etranger_orders = df.loc[mask_etranger, 'Order ID'].unique().tolist()
    
    is_all_zebra = df.groupby('Order ID')['Catégorie'].transform(lambda x: (x == 'Zebra').all())
    zebra_orders = df.loc[is_all_zebra, 'Order ID'].unique().tolist()
    
    commandes_sans_prix = set(etranger_orders + zebra_orders)
    df.loc[df['Order ID'].isin(commandes_sans_prix), 'Total Order Price'] = pd.NA
    
    mask_doublons = df.duplicated(subset=['Order ID'], keep='first')
    df.loc[mask_doublons, 'Total Order Price'] = pd.NA
    df.loc[mask_doublons, 'Catégorie'] = pd.NA
    
    nb_commandes_uniques = df['Order ID'].nunique()
    nb_manuelles = len(commandes_sans_prix)

    temp_buffer = io.BytesIO()
    df.to_excel(temp_buffer, index=False)
    temp_buffer.seek(0)
    
    wb = openpyxl.load_workbook(temp_buffer)
    ws = wb.active
    ws.title = "Facturation WebToPrint"
    
    header_fill, orange_fill = PatternFill("solid", fgColor="183153"), PatternFill("solid", fgColor="FFC000")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.
