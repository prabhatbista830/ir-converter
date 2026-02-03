import streamlit as st
st.write("hello - The app is RUnning")
import pandas as pd
from openpyxl import load_workbook
import os
import re

# --- CONFIGURATION ---
INPUT_CMM_FILE = "CMM_Result.xlsx"
TEMPLATE_IR_FILE = "IR_Template.xlsx"
OUTPUT_FILE = "Final_Report_Done.xlsx"

CMM_CHAR_COL = "Characteristic" 
CMM_VAL_COL  = "Actual"         

IR_ID_COL = "5. Char No."
IR_RES_COL = "9. Results"

def extract_base_number(text):
    if pd.isna(text): return None
    match = re.search(r'(\d+)', str(text))
    return match.group(1) if match else None

def format_val(val):
    """Ensures exactly 4 digits after the decimal point."""
    try:
        return f"{float(val):.4f}"
    except:
        return str(val)

def main():
    print("--- Running Precision-4 Automation ---")

    if not os.path.exists(INPUT_CMM_FILE):
        print(f"ERROR: File {INPUT_CMM_FILE} not found.")
        return

    # 1. READ CMM DATA
    df_raw = pd.read_excel(INPUT_CMM_FILE, header=None, nrows=50)
    header_row_idx = next((i for i, row in df_raw.iterrows() if row.astype(str).str.contains(CMM_CHAR_COL, case=False).any()), None)
    
    if header_row_idx is None:
        print(f"ERROR: Could not find header '{CMM_CHAR_COL}'.")
        return

    df_cmm = pd.read_excel(INPUT_CMM_FILE, header=header_row_idx)
    df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]

    char_col_name = next((c for c in df_cmm.columns if CMM_CHAR_COL.upper() in c), None)
    val_col_name = next((c for c in df_cmm.columns if CMM_VAL_COL.upper() in c), None)

    # 2. GROUPING DATA
    cmm_results = {}

    for _, row in df_cmm.iterrows():
        raw_text = str(row[char_col_name]).strip().upper()
        base_num = extract_base_number(raw_text)
        if not base_num: continue

        # COORDINATE GUARD: Skip X, Y, Z components
        is_coordinate = any(raw_text.endswith(f".{c}") or 
                            raw_text.endswith(f" {c}") or 
                            raw_text == c for c in ['X', 'Y', 'Z'])
        if is_coordinate: continue

        try:
            val = float(row[val_col_name])
        except:
            continue

        if base_num not in cmm_results:
            cmm_results[base_num] = {'master': None, 'samples': []}

        if raw_text == base_num or raw_text == f"{base_num}.0":
            cmm_results[base_num]['master'] = val
        else:
            cmm_results[base_num]['samples'].append(val)

    # 3. FILLING IR TEMPLATE
    wb = load_workbook(TEMPLATE_IR_FILE)
    ws = wb.active

    id_col_idx, res_col_idx, start_row = None, None, None
    for r in range(1, 30):
        for c in range(1, ws.max_column + 1):
            cell_val = str(ws.cell(row=r, column=c).value)
            if IR_ID_COL in cell_val: 
                id_col_idx, start_row = c, r
            if IR_RES_COL in cell_val: 
                res_col_idx = c

    # 4. WRITE THE DATA WITH 4-DECIMAL FORMATTING
    for r in range(start_row + 1, ws.max_row + 1):
        ir_id = extract_base_number(ws.cell(row=r, column=id_col_idx).value)
        
        if ir_id in cmm_results:
            data = cmm_results[ir_id]
            
            if data['master'] is not None:
                final_output = format_val(data['master'])
            elif data['samples']:
                vals = data['samples']
                if len(vals) > 1 and min(vals) != max(vals):
                    # Formats both sides of the range to 4 decimal places
                    final_output = f"{format_val(min(vals))} - {format_val(max(vals))}"
                else:
                    final_output = format_val(vals[0])
            else:
                continue
            
            ws.cell(row=r, column=res_col_idx).value = final_output

    wb.save(OUTPUT_FILE)
    print(f"--- Process Complete! All values formatted to 4 decimals. ---")

if __name__ == "__main__":
    main()
