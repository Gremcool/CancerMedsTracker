import pandas as pd
import re
from database import get_connection

def extract_base_drug_name(full_name):
    """Isolates parent molecule definitions from raw input descriptions."""
    if pd.isna(full_name):
        return "UNKNOWN"
    
    text = str(full_name).strip()
    match = re.search(r'\d+', text)
    if match:
        base = text[:match.start()].strip()
        base = re.sub(r'[\s,波动\-]+$', '', base)
        if base:
            return base.upper()
            
    words = text.split()
    return words[0].upper() if words else "UNKNOWN"

def load_excel_sheets(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    return xl.sheet_names

def load_excel(uploaded_file, sheet_name=0):
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)

def prepare_medicines(df):
    """Maps spreadsheet columns into the generic layout directory data frames safely."""
    desc_col = next((c for c in df.columns if 'DESCRIPTION' in str(c).upper() or 'ITEM' in str(c).upper()), None)
    
    if not desc_col:
        return pd.DataFrame()
        
    df_clean = df[[desc_col]].dropna()
    df_clean.columns = ['medicine_name']
    df_clean['medicine_name'] = df_clean['medicine_name'].astype(str).str.strip()
    
    df_clean = df_clean[df_clean['medicine_name'] != '']
    df_clean = df_clean[~df_clean['medicine_name'].str.contains("COMPATIBILITY REPORT|RUN ON|TOTAL", case=False, na=False)]
    
    df_clean['base_drug_name'] = df_clean['medicine_name'].apply(extract_base_drug_name)
    df_clean = df_clean[df_clean['base_drug_name'] != "UNKNOWN"]
    
    return df_clean.drop_duplicates(subset=['medicine_name'])

def save_medicines(df):
    """Inserts lines directly into database tracking arrays."""
    conn = get_connection()
    cur = conn.cursor()
    for _, row in df.iterrows():
        cur.execute("""
        INSERT OR IGNORE INTO medicines (base_drug_name, medicine_name)
        VALUES (?, ?)
        """, (row['base_drug_name'], row['medicine_name']))
    conn.commit()
    conn.close()