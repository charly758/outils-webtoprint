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
    df.columns = df.columns.str.strip() 
    colonnes_requises = ['Customer Name', 'Order ID', 'Product Name', 'Total Quantity', 'DateOrderCreated', 'Ship_Company', 'Ship_City', 'Ship Country Name', 'Total Order Price']
    
    if 'Product Name' not in df.columns: raise ValueError("Le fichier ne contient pas la colonne 'Product Name'.")

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
    return final_buffer

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
    
    temp_buffer = io.BytesIO()
    df.to_excel(temp_buffer, index=False)
    temp_buffer.seek(0)
    
    wb = openpyxl.load_workbook(temp_buffer)
    ws = wb.active
    ws.title = "Etat des Stocks"
    
    header_fill = PatternFill("solid", fgColor="183153")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")
        
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
        
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 65
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 20
    ws.auto_filter.ref = ws.dimensions

    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer

# ==========================================
# 3. FONCTION : SUIVI COLIS (DPD)
# ==========================================
def traiter_suivi_colis(file, reference_dossier):
    df = pd.read_excel(file)
    wb_in = openpyxl.load_workbook(file)
    ws_in = wb_in.active
    
    urls = []
    for r in range(2, len(df) + 2):
        cell = ws_in.cell(row=r, column=24)
        urls.append(cell.hyperlink.target if cell.hyperlink else cell.value)
        
    if len(df.columns) >= 24:
        df.iloc[:, 23] = urls 
    
    ref_col_name = df.columns[1]
    df[ref_col_name] = df[ref_col_name].astype(str).str.strip()
    reference_propre = str(reference_dossier).strip()
    
    df_filtre = df[df[ref_col_name] == reference_propre].copy()
    
    if df_filtre.empty:
        raise ValueError(f"⚠️ Aucune commande trouvée pour le dossier N° {reference_propre}. Vérifiez le numéro saisi.")
    
    indices_a_garder = [0, 1, 6, 9, 11, 13, 14, 15, 17, 18, 21, 22, 23]
    
    if len(df.columns) <= max(indices_a_garder):
        raise ValueError("Le fichier Excel importé n'a pas le format attendu (pas assez de colonnes).")
        
    df_final = df_filtre.iloc[:, indices_a_garder].copy()
    
    for col in df_final.columns:
        if pd.api.types.is_datetime64_any_dtype(df_final[col]):
            df_final[col] = df_final[col].dt.strftime('%d/%m/%Y')
            
    temp_buffer = io.BytesIO()
    df_final.to_excel(temp_buffer, index=False)
    temp_buffer.seek(0)
    
    wb = openpyxl.load_workbook(temp_buffer)
    ws = wb.active
    ws.title = f"Dossier {reference_propre}"
    
    header_fill = PatternFill("solid", fgColor="183153")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")
        
    for r in range(2, ws.max_row + 1):
        cell_lien = ws.cell(row=r, column=13) 
        val = str(cell_lien.value)
        if val.startswith("http"):
            cell_lien.hyperlink = val
            cell_lien.value = "Voir le suivi" 
            cell_lien.font = Font(color="0563C1", underline="single") 
            cell_lien.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_length = max((len(str(c.value)) if c.value else 0 for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_length + 3, 50)
        
    ws.auto_filter.ref = ws.dimensions

    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer

# ==========================================
# BARRE LATÉRALE (MENU)
# ==========================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3714/3714856.png", width=60) # Petite icône sympa
st.sidebar.markdown("### Votre boîte à outils")
st.sidebar.markdown("Sélectionnez le traitement que vous souhaitez effectuer :")

outil = st.sidebar.radio("", [
    "📊 Facturation Mensuelle", 
    "📦 Nettoyage des Stocks",
    "🚚 Extraction Suivi Colis (DPD)"
])

st.sidebar.markdown("---")
st.sidebar.caption("🔧 Développé sur-mesure pour vos opérations Web-to-Print.")


# ==========================================
# CORPS DE L'APPLICATION
# ==========================================
st.markdown('<p class="main-title">Portail Web-to-Print</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Automatisation et traitement de vos exports Excel</p>', unsafe_allow_html=True)

if outil == "📊 Facturation Mensuelle":
    st.header("Facturation Mensuelle")
    st.info("💡 **Instructions :** Glissez l'export brut des commandes (ex: `2.xlsx`) pour générer le tableau classé, fusionné et valorisé.")
    
    fichier = st.file_uploader("📂 Importez le fichier Excel de facturation", type=["xlsx"])
    if fichier:
        try:
            with st.spinner("⏳ Traitement en cours..."):
                fichier_final = traiter_facturation(fichier)
            st.success("✅ Tableau généré avec succès ! Les doublons sont fusionnés et l'export est catégorisé.")
            st.download_button(
                label="📥 Télécharger le tableau final", 
                data=fichier_final, 
                file_name="Facturation_Finale.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" # Rend le bouton bien visible
            )
        except Exception as e:
            st.error(f"❌ Une erreur est survenue : {e}")

elif outil == "📦 Nettoyage des Stocks":
    st.header("Nettoyage du Rapport de Stock")
    st.info("💡 **Instructions :** Glissez l'export brut des stocks (ex: `Report.xlsx`) pour générer la version épurée et formatée.")
    
    fichier = st.file_uploader("📂 Importez le rapport de stock", type=["xlsx"])
    if fichier:
        try:
            with st.spinner("⏳ Traitement en cours..."):
                fichier_final = traiter_stock(fichier)
            st.success("✅ Stock nettoyé avec succès ! Formatage et filtres appliqués.")
            st.download_button(
                label="📥 Télécharger l'état des stocks", 
                data=fichier_final, 
                file_name="Stock_Final.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        except Exception as e:
            st.error(f"❌ Une erreur est survenue : {e}")

elif outil == "🚚 Extraction Suivi Colis (DPD)":
    st.header("Extraction des Suivis Colis (DPD)")
    st.info("💡 **Instructions :** Glissez votre tableau de suivi DPD brut puis saisissez le numéro de dossier. L'outil génèrera un tableau ciblé avec des liens de suivi cliquables.")
    
    fichier = st.file_uploader("📂 Importez le fichier de suivi", type=["xlsx"])
    
    if fichier:
        numero_dossier = st.text_input("🎯 Saisissez le numéro de dossier (Votre référence 1) :", placeholder="ex: 2012217")
        
        if st.button("Rechercher le dossier", type="primary"):
            if numero_dossier:
                try:
                    with st.spinner(f"⏳ Recherche du dossier {numero_dossier} en cours..."):
                        fichier_traite = traiter_suivi_colis(fichier, numero_dossier)
                        
                        st.session_state['fichier_dpd_pret'] = fichier_traite
                        st.session_state['dossier_dpd_actuel'] = numero_dossier
                except Exception as e:
                    st.error(f"❌ {e}")
            else:
                st.warning("⚠️ Veuillez saisir un numéro de dossier avant de cliquer sur Rechercher.")
                
    if 'fichier_dpd_pret' in st.session_state and 'dossier_dpd_actuel' in st.session_state:
        st.success(f"✅ Dossier {st.session_state['dossier_dpd_actuel']} extrait avec succès !")
        st.download_button(
            label=f"📥 Télécharger le suivi ({st.session_state['dossier_dpd_actuel']})", 
            data=st.session_state['fichier_dpd_pret'], 
            file_name=f"Suivi_Dossier_{st.session_state['dossier_dpd_actuel']}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
