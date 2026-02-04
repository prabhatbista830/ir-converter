import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import re
import io

# --- LOGIC FUNCTIONS (IR CONVERTER - UNCHANGED) ---
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

def is_coordinate_basic(char_name):
    name = str(char_name).strip().upper()
    return any(name.endswith(f".{c}") or name.endswith(f" {c}") or name == c for c in ['X', 'Y', 'Z'])

# --- THE WEBSITE INTERFACE ---
st.set_page_config(page_title="CMM Quality Suite", page_icon="📊", layout="wide")

# --- NAVIGATION ---
page = st.sidebar.radio("Navigation Menu", ["🏠 Home", "📝 IR Converter", "⚠️ Discrepancy Report"])

# --- PAGE 1: HOME ---
if page == "🏠 Home":
    st.title("📊 CMM Quality Suite")
    st.write("Welcome! Use the sidebar to switch between tools.")
    st.info("IR Converter: Verified logic. Discrepancy Report: Multi-file batching with 4-decimal positive numbers.")

# --- PAGE 2: IR CONVERTER (STABLE VERSION) ---
elif page == "📝 IR Converter":
    st.title("📊 CMM Result to IR Automator")
    uploaded_cmm = st.file_uploader("Upload CMM Result", type=["xlsx"], key="ir_cmm_up")
    uploaded_template = st.file_uploader("Upload IR Template", type=["xlsx"], key="ir_tmp_up")

    if uploaded_cmm and uploaded_template:
        if st.button("🚀 Process and Generate Report"):
            try:
                df_raw = pd.read_excel(uploaded_cmm, header=None, nrows=50)
                header_row_idx = next((i for i, row in df_raw.iterrows() if row.astype(str).str.contains("Characteristic", case=False).any()), None)
                if header_row_idx is None:
                    st.error("Could not find 'Characteristic' column.")
                else:
                    df_cmm = pd.read_excel(uploaded_cmm, header=header_row_idx)
                    df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]
                    cmm_results = {}
                    for _, row in df_cmm.iterrows():
                        raw_text = str(row.get("CHARACTERISTIC", "")).strip().upper()
                        base_num = extract_base_number(raw_text)
                        if not base_num or is_coordinate_basic(raw_text): continue
                        try: val = float(row.get("ACTUAL", 0))
                        except: continue
                        if base_num not in cmm_results: cmm_results[base_num] = {'master': None, 'samples': []}
                        if raw_text == base_num or raw_text == f"{base_num}.0": cmm_results[base_num]['master'] = val
                        else: cmm_results[base_num]['samples'].append(val)
                    
                    wb = load_workbook(uploaded_template)
                    ws = wb.active
                    id_col_idx, res_col_idx, start_row = None, None, None
                    for r in range(1, 31):
                        for c in range(1, ws.max_column + 1):
                            cell_val = str(ws.cell(row=r, column=c).value)
                            if "5. Char No." in cell_val: id_col_idx, start_row = c, r
                            if "9. Results" in cell_val: res_col_idx = c
                    
                    if id_col_idx and res_col_idx:
                        for r in range(start_row + 1, ws.max_row + 1):
                            ir_id = extract_base_number(ws.cell(row=r, column=id_col_idx).value)
                            if ir_id in cmm_results:
                                data = cmm_results[ir_id]
                                if data['master'] is not None: final_output = format_val(data['master'])
                                elif data['samples']:
                                    vals = data['samples']
                                    final_output = f"{format_val(min(vals))} - {format_val(max(vals))}" if len(vals) > 1 and min(vals) != max(vals) else format_val(vals[0])
                                else: continue
                                ws.cell(row=r, column=res_col_idx).value = final_output
                        output = io.BytesIO()
                        wb.save(output)
                        st.success("✅ IR Report Generated!")
                        st.download_button("📥 Download Final IR", output.getvalue(), "Final_Report_Done.xlsx")
            except Exception as e: st.error(f"Error: {e}")

# --- PAGE 3: DISCREPANCY REPORT (4 DECIMALS & BLANKS) ---
elif page == "⚠️ Discrepancy Report":
    st.title("⚠️ Batch Out-of-Tolerance Reporter")
    st.write("Upload multiple CMM files. Failures show as positive numbers with 4 decimals; passes remain blank.")
    
    uploaded_files = st.file_uploader("Upload CMM Results (Multi-select)", type=["xlsx"], accept_multiple_files=True, key="oot_batch")

    if uploaded_files:
        if st.button("🔍 Generate Combined Discrepancy Report"):
            all_part_data = [] 
            
            try:
                for uploaded_file in uploaded_files:
                    # 1. SN from F8
                    df_sn = pd.read_excel(uploaded_file, header=None, nrows=10, usecols="F")
                    sn_val = df_sn.iloc[7, 0]
                    
                    # 2. Data Load
                    df_raw_oot = pd.read_excel(uploaded_file, header=None, nrows=50)
                    h_idx = next((i for i, row in df_raw_oot.iterrows() if row.astype(str).str.contains("Characteristic", case=False).any()), 12)
                    df_data = pd.read_excel(uploaded_file, header=h_idx)
                    df_data.columns = [str(c).strip().upper() for c in df_data.columns]

                    part_row = {"SN": sn_val}

                    for _, row in df_data.iterrows():
                        name = str(row.get("CHARACTERISTIC", "")).strip()
                        if is_coordinate_basic(name) or name.lower() == "nan" or name == "": continue
                        
                        try:
                            act, nom = float(row['ACTUAL']), float(row['NOMINAL'])
                            u_tol, l_tol = float(row['UPPER TOL']), float(row['LOWER TOL'])
                            
                            if act > (nom + u_tol) or act < (nom + l_tol):
                                if l_tol == 0: t_str = f"+ {abs(u_tol)}"
                                elif u_tol == 0: t_str = f"- {abs(l_tol)}"
                                else: t_str = f"+/- {abs(u_tol)}"
                                
                                col_header = f"Dim#{name} ({nom} {t_str})"
                                # Force positive and exactly 4 decimals
                                part_row[col_header] = f"{abs(act):.4f}"
                        except: continue
                    
                    all_part_data.append(part_row)

                if all_part_data:
                    final_oot_df = pd.DataFrame(all_part_data)
                    final_oot_df = final_oot_df.fillna("")
                    
                    st.write("### Combined Failure Summary:")
                    st.dataframe(final_oot_df)
                    
                    out_oot = io.BytesIO()
                    with pd.ExcelWriter(out_oot, engine='xlsxwriter') as writer:
                        final_oot_df.to_excel(writer, index=False)
                    st.download_button("📥 Download Combined Discrepancies", out_oot.getvalue(), "Combined_Discrepancy_Report.xlsx")
                else:
                    st.success("✅ No discrepancies found!")
            except Exception as e:
                st.error(f"Error processing files: {e}")
