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
        .main-title { font-size: 2.5rem; font-weight: 700; color: #222222; margin-bottom: 0px; }
        .subtitle { font-size: 1.1rem; color: #BA2025; margin-bottom: 30px; font-weight: 600; }
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
        cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")
    
    col_order, col_price, col_cat, col_country = None, None, None, None
    for idx, cell in enumerate(ws[1], 1):
        if cell.value == 'Order ID': col_order = idx
        elif cell.value == 'Total Order Price': col_price = idx
        elif cell.value == 'Catégorie': col_cat = idx
        elif cell.value == 'Ship Country Name': col_country = idx
        
    for r in range(2, ws.max_row + 1):
        order_id = ws.cell(row=r, column=col_order).value if col_order else None
        if order_id in commandes_sans_prix:
            if col_price: ws.cell(row=r, column=col_price).fill = orange_fill
            if col_country: ws.cell(row=r, column=col_country).fill = orange_fill

        if col_price:
            c_price = ws.cell(row=r, column=col_price)
            if c_price.value is not None and str(c_price.value).strip() != "":
                try: c_price.value = float(c_price.value)
                except: pass
                c_price.number_format = '#,##0.00 €'
            c_price.alignment = Alignment(horizontal="center", vertical="center")
            
        if col_cat: ws.cell(row=r, column=col_cat).alignment = Alignment(horizontal="center", vertical="center")
            
    for col in ws.columns:
        max_length = max((len(f"{c.value} €" if c.number_format == '#,##0.00 €' and c.value else str(c.value) if c.value else "") for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3
        
    start_row = 2
    current_order = ws.cell(row=2, column=col_order).value if col_order else None
    
    for r in range(3, ws.max_row + 2):
        order = ws.cell(row=r, column=col_order).value if r <= ws.max_row and col_order else None
        if order != current_order:
            if r - 1 > start_row:
                if col_price: ws.merge_cells(start_row=start_row, start_column=col_price, end_row=r-1, end_column=col_price)
                if col_cat: ws.merge_cells(start_row=start_row, start_column=col_cat, end_row=r-1, end_column=col_cat)
            current_order, start_row = order, r

    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer, nb_commandes_uniques, nb_manuelles

# ==========================================
# 2. FONCTION : STOCK
# ==========================================
def traiter_stock(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip() 
    colonnes_requises = ['N° de produit', 'Nom de produit', 'Stock réel', 'Unité de mesure']
    
    missing_cols = [c for c in colonnes_requises if c not in df.columns]
    if missing_cols: raise ValueError(f"Il manque des colonnes : {', '.join(missing_cols)}")
        
    df = df[colonnes_requises].copy()
    nb_articles = len(df)
    
    temp_buffer = io.BytesIO()
    df.to_excel(temp_buffer, index=False)
    temp_buffer.seek(0)
    
    wb = openpyxl.load_workbook(temp_buffer)
    ws = wb.active
    ws.title = "Etat des Stocks"
    
    header_fill = PatternFill("solid", fgColor="183153")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
