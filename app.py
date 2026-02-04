import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import re
import io

# --- LOGIC FUNCTIONS ---
def extract_base_number(text):
    if pd.isna(text): 
        return None
    match = re.search(r'(\d+)', str(text))
    return match.group(1) if match else None

def format_val(val):
    try:
        return f"{float(val):.4f}"
    except:
        return str(val)

# --- THE WEBSITE INTERFACE ---
st.set_page_config(page_title="CMM to IR Converter", page_icon="📊")

st.title("📊 CMM Result to IR Automator")
st.write("Upload your CMM Excel export and your IR Template to automate the data entry.")

# 1. FILE UPLOADERS
uploaded_cmm = st.file_uploader("Step 1: Upload CMM Result (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Step 2: Upload IR Template (Excel)", type=["xlsx"])

if uploaded_cmm and uploaded_template:
    if st.button("🚀 Process and Generate Report"):
        with st.spinner("Processing data..."):
            try:
                # --- READ CMM DATA ---
                # Read first 50 rows to find where the actual data starts
                df_raw = pd.read_excel(uploaded_cmm, header=None, nrows=50)
                header_row_idx = next((i for i, row in df_raw.iterrows() if row.astype(str).str.contains("Characteristic", case=False).any()), None)
                
                if header_row_idx is None:
                    st.error("Could not find 'Characteristic' column in CMM file. Please check the file format.")
                else:
                    df_cmm = pd.read_excel(uploaded_cmm, header=header_row_idx)
                    df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]
                    
                    cmm_results = {}
                    for _, row in df_cmm.iterrows():
                        raw_text = str(row.get("CHARACTERISTIC", "")).strip().upper()
                        base_num = extract_base_number(raw_text)
                        
                        if not base_num: 
                            continue
                        
                        # Filter out coordinate noise
                        is_coordinate = any(raw_text.endswith(f".{c}") or raw_text.endswith(f" {c}") or raw_text == c for c in ['X', 'Y', 'Z'])
                        if is_coordinate: 
                            continue

                        try:
                            val = float(row.get("ACTUAL", 0))
                        except:
                            continue

                        if base_num not in cmm_results:
                            cmm_results[base_num] = {'master': None, 'samples': []}

                        # Check if it's the primary characteristic or a sub-sample
                        if raw_text == base_num or raw_text == f"{base_num}.0":
                            cmm_results[base_num]['master'] = val
                        else:
                            cmm_results[base_num]['samples'].append(val)

                    # --- FILL TEMPLATE ---
                    wb = load_workbook(uploaded_template)
                    ws = wb.active
                    
                    id_col_idx, res_col_idx, start_row = None, None, None
                    
                    # Search for headers in the template (top 30 rows)
                    for r in range(1, 31):
                        for c in range(1, ws.max_column + 1):
                            cell_val = str(ws.cell(row=r, column=c).value)
                            if "5. Char No." in cell_val: 
                                id_col_idx, start_row = c, r
                            if "9. Results" in cell_val: 
                                res_col_idx = c

                    if id_col_idx is None or res_col_idx is None:
                        st.error("Could not find '5. Char No.' or '9. Results' in the Template. Please check headers.")
                    else:
                        # Iterate through the template rows and fill data
                        for r in range(start_row + 1, ws.max_row + 1):
                            raw_id_val = ws.cell(row=r, column=id_col_idx).value
                            ir_id = extract_base_number(raw_id_val)
                            
                            if ir_id in cmm_results:
                                data = cmm_results[ir_id]
                                if data['master'] is not None:
                                    final_output = format_val(data['master'])
                                elif data['samples']:
                                    vals = data['samples']
                                    # If multiple samples exist, show the range (Min - Max)
                                    if len(vals) > 1 and min(vals) != max(vals):
                                        final_output = f"{format_val(min(vals))} - {format_val(max(vals))}"
                                    else:
                                        final_output = format_val(vals[0])
                                else: 
                                    continue
                                
                                ws.cell(row=r, column=res_col_idx).value = final_output

                        # --- SAVE TO MEMORY FOR DOWNLOAD ---
                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)
                        
                        st.success("✅ Report Generated Successfully!")
                        st.download_button(
                            label="📥 Download Final IR Report",
                            data=output,
                            file_name="Final_Report_Done.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")
