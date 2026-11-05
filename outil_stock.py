import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
import io

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
    return final_buffer, nb_articles
