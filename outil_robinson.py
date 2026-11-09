import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter
import io

def traiter_robinson(file):
    # 1. Lecture du fichier texte
    df = pd.read_csv(file, sep='\t')
    
    # 2. Suppression de la colonne RFA
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
    
    # 5. Création du fichier Excel
    temp_buffer = io.BytesIO()
    df.to_excel(temp_buffer, index=False, header=False)
    temp_buffer.seek(0)
    
    wb = openpyxl.load_workbook(temp_buffer)
    ws = wb.active
    ws.title = "Feuille 1 - Robinson"
    
    ws.insert_rows(1, 2)
    
    ws.merge_cells('A1:F1')
    ws['A1'] = "Robinson"
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws['A1'].font = Font(size=14, bold=True)
    
    headers = ['Contact', 'Date de facture', 'N° de facture', 'Montant (hors TVA)', 'Commission', 'Montant Commission']
    header_fill = PatternFill("solid", fgColor="009CA6") 
    header_font = Font(bold=True, color="FFFFFF")
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # 6. Formatage des cellules avec VERROUILLAGE DE L'EURO
    for r in range(3, ws.max_row + 1):
        ws.cell(row=r, column=3).number_format = '0'
        ws.cell(row=r, column=4).number_format = '#,##0.00" "€'  # Euro forcé
        ws.cell(row=r, column=5).number_format = '0%'
        ws.cell(row=r, column=6).number_format = '#,##0.00" "€'  # Euro forcé
        
    # ==========================================
    # LIGNE DES TOTAUX
    # ==========================================
    total_row = ws.max_row + 1
    
    cell_total_label = ws.cell(row=total_row, column=3, value="TOTAL")
    cell_total_label.font = Font(bold=True)
    cell_total_label.alignment = Alignment(horizontal="right")
    
    cell_sum_montant = ws.cell(row=total_row, column=4, value=f"=SUM(D3:D{total_row-1})")
    cell_sum_montant.font = Font(bold=True)
    cell_sum_montant.number_format = '#,##0.00" "€'  # Euro forcé
    
    cell_sum_comm = ws.cell(row=total_row, column=6, value=f"=SUM(F3:F{total_row-1})")
    cell_sum_comm.font = Font(bold=True)
    cell_sum_comm.number_format = '#,##0.00" "€'  # Euro forcé
    # ==========================================
        
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_length = max((len(str(ws.cell(row=r, column=col_idx).value)) if ws.cell(row=r, column=col_idx).value else 0 for r in range(2, ws.max_row + 1)), default=0)
        ws.column_dimensions[col_letter].width = max_length + 3

    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer, nb_lignes
