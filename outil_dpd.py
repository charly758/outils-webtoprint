import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
import io

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
        raise ValueError(f"Aucun colis trouvé pour le dossier N° {reference_propre}.")
    
    indices_a_garder = [0, 1, 6, 9, 11, 13, 14, 15, 17, 18, 21, 22, 23]
    
    if len(df.columns) <= max(indices_a_garder):
        raise ValueError("Le fichier ne contient pas les colonnes DPD habituelles.")
        
    df_final = df_filtre.iloc[:, indices_a_garder].copy()
    nb_colis = len(df_final)
    
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
    return final_buffer, nb_colis
