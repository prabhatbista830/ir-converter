import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import re
import io

# --- UNIVERSAL SETTINGS ---
st.set_page_config(page_title="CMM Quality Suite", page_icon="📊", layout="wide")

# --- ORIGINAL LOGIC FUNCTIONS ---
def extract_base_number(text):
    if pd.isna(text): return None
    match = re.search(r'(\d+)', str(text))
    return match.group(1) if match else None

def format_val(val):
    try:
        return f"{float(val):.4f}"
    except:
        return str(val)

def is_coordinate(char_name):
    name = str(char_name).strip().upper()
    suffixes = ('.X', '.Y', '.Z', '.A', '.B', '.C', ' X', ' Y', ' Z')
    return name.endswith(suffixes) or name in ['X', 'Y', 'Z', 'A', 'B', 'C']

# --- NAVIGATION ---
page = st.sidebar.radio("Navigation Menu", ["🏠 Home", "📝 IR Converter", "⚠️ Discrepancy Feature"])

# --- PAGE 1: HOME ---
if page == "🏠 Home":
    st.title("📊 CMM Quality Suite")
    st.write("Welcome! This tool handles your quality reporting automation.")
    st.info("Select 'IR Converter' for the original template logic or 'Discrepancy Feature' for failure reports.")

# --- PAGE 2: IR CONVERTER (YOUR ORIGINAL CODE INTEGRATED) ---
elif page == "📝 IR Converter":
    st.title("📝 IR Template Automator")
    uploaded_cmm = st.file_uploader("Step 1: Upload CMM Result (Excel)", type=["xlsx"], key="ir_cmm")
    uploaded_template = st.file_uploader("Step 2: Upload IR Template (Excel)", type=["xlsx"], key="ir_tmp")

    if uploaded_cmm and uploaded_template:
        if st.button("🚀 Process and Generate Report"):
            with st.spinner("Processing data..."):
                try:
                    # --- READ CMM DATA ---
                    df_raw = pd.read_excel(uploaded_cmm, header=None, nrows=50)
                    header_row_idx = next((i for i, row in df_raw.iterrows() if row.astype(str).str.contains("Characteristic", case=False).any()), None)
                    
                    if header_row_idx is None:
                        st.error("Could not find 'Characteristic' column in CMM file.")
                    else:
                        df_cmm = pd.read_excel(uploaded_cmm, header=header_row_idx)
                        df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]
                        
                        cmm_results = {}
                        for _, row in df_cmm.iterrows():
                            raw_text = str(row.get("CHARACTERISTIC", "")).strip().upper()
                            base_num = extract_base_number(raw_text)
                            if not base_num: continue
                            
                            if is_coordinate(raw_text): continue

                            try:
                                val = float(row.get("ACTUAL", 0))
                            except: continue

                            if base_num not in cmm_results:
                                cmm_results[base_num] = {'master': None, 'samples': []}

                            # Logic for Master vs Samples
                            if raw_text == base_num or raw_text == f"{base_num}.0":
                                cmm_results[base_num]['master'] = val
                            else:
                                cmm_results[base_num]['samples'].append(val)

                        # --- FILL TEMPLATE ---
                        wb = load_workbook(uploaded_template)
                        ws = wb.active
                        
                        id_col_idx, res_col_idx, start_row = None, None, None
                        for r in range(1, 30):
                            for c in range(1, ws.max_column + 1):
                                cell_val = str(ws.cell(row=r, column=c).value)
                                if "5. Char No." in cell_val: id_col_idx, start_row = c, r
                                if "9. Results" in cell_val: res_col_idx = c

                        if id_col_idx and res_col_idx:
                            for r in range(start_row + 1, ws.max_row + 1):
                                ir_id = extract_base_number(ws.cell(row=r, column=id_col_idx).value)
                                if ir_id in cmm_results:
                                    data = cmm_results[ir_id]
                                    if data['master'] is not None:
                                        final_output = format_val(data['master'])
                                    elif data['samples']:
                                        vals = data['samples']
                                        final_output = f"{format_val(min(vals))} - {format_val(max(vals))}" if len(vals) > 1 and min(vals) != max(vals) else format_val(vals[0])
                                    else: continue
                                    ws.cell(row=r, column=res_col_idx).value = final_output

                            output = io.BytesIO()
                            wb.save(output)
                            st.success("✅ Report Generated Successfully!")
                            st.download_button("📥 Download Final IR Report", output.getvalue(), "Final_Report_Done.xlsx")
                except Exception as e:
                    st.error(f"An error occurred: {e}")

# --- PAGE 3: DISCREPANCY FEATURE ---
elif page == "⚠️ Discrepancy Feature":
    st.title("⚠️ Out-of-Tolerance Reporter")
    uploaded_oot = st.file_uploader("Upload CMM Result", type=["xlsx"], key="oot_cmm")

    if uploaded_oot:
        if st.button("🔍 Generate Discrepancy Report"):
            try:
                # 1. SN from F8
                df_sn = pd.read_excel(uploaded_oot, header=None, nrows=10, usecols="F")
                sn_val = df_sn.iloc[7, 0]
                
                #
