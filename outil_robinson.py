import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter
import io

def traiter_robinson(file):
    # 1. Lecture du fichier texte (gère les tabulations de l'export ERP)
    df = pd.read_csv(file, sep='\t')
    
    # 2. Suppression de la colonne RFA si elle existe
    if 'RFA' in df.columns:
        df = df.drop(columns=['RFA'])
        
    # 3. Nettoyage des données pour Excel
    if 'N° de facture' in df.columns:
        df['N° de facture'] = df['N° de facture'].astype(str).str.replace(r'\s+', '', regex=True).astype(int)
        
    if 'Montant (hors TVA)' in df.columns:
        df['Montant (hors TVA)'] = df['Montant (hors TVA)'].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.').astype(float)
        
    if 'Commission' in df.columns:
        df['Commission'] = df['Commission'].astype(str).str.replace('%', '', regex=False).astype(float) / 100.0
        
    # 4. Calcul du montant final
    df['Montant Commission'] = df['Montant (hors TVA)'] * df['Commission']
    nb_lignes = len(df)
    
    # 5. Création du fichier Excel sur-mesure
    temp_buffer = io.BytesIO()
    df.to_excel(temp_buffer, index=False, header=False)
    temp_buffer.seek(0)
    
    wb = openpyxl.load_workbook(temp_buffer)
    ws = wb.active
    ws.title = "Feuille 1 - Robinson"
    
    # On insère 2 lignes vierges au début
    ws.insert_rows(1, 2)
    
    # Ligne 1 : Grand Titre centré
    ws.merge_cells('A1:F1')
    ws['A1'] = "Robinson"
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws['A1'].font = Font(size=14, bold=True)
    
    # Ligne 2 : En-têtes aux couleurs MicroLynx
    headers = ['Contact', 'Date de facture', 'N° de facture', 'Montant (hors TVA)', 'Commission', 'Montant Commission']
    header_fill = PatternFill("solid", fgColor="009CA6") 
    header_font = Font(bold=True, color="FFFFFF")
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # 6. Formatage des cellules (monétaires et pourcentages purs)
    for r in range(3, ws.max_row + 1):
        ws.cell(row=r, column=3).number_format = '0'           # N° Facture
        ws.cell(row=r, column=4).number_format = '#,##0.00 €'  # Montant
        ws.cell(row=r, column=5).number_format = '0%'          # Commission
        ws.cell(row=r, column=6).number_format = '#,##0.00 €'  # Montant calculé
        
    # 7. Ajustement des largeurs de colonnes (Corrigé pour ignorer la ligne 1 fusionnée)
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        # On calcule la largeur max en commençant à la ligne 2 pour éviter le bug "MergedCell"
        max_length = max((len(str(ws.cell(row=r, column=col_idx).value)) if ws.cell(row=r, column=col_idx).value else 0 for r in range(2, ws.max_row + 1)), default=0)
        ws.column_dimensions[col_letter].width = max_length + 3

    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer, nb_lignes
