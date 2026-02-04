import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import re
import io

st.set_page_config(page_title="CMM Quality Suite", layout="wide")

# --- HELPER FUNCTIONS ---
def get_clean_id(text):
    """Removes MAX/MIN/XYZ but keeps decimals like 30.1"""
    if pd.isna(text): return None
    s = str(text).strip().upper()
    # Remove text suffixes like MAX, MIN, etc.
    s = re.sub(r'(MAX|MIN|MIN-MAX|\^MIN|\^MAX|POS|NEG).*', '', s).strip()
    return s

def extract_numbers_only(text):
    if pd.isna(text): return None
    match = re.search(r'(\d+(\.\d+)?)', str(text))
    return match.group(1) if match else None

def is_coordinate(char_name):
    name = str(char_name).strip().upper()
    suffixes = ('.X', '.Y', '.Z', '.A', '.B', '.C', ' X', ' Y', ' Z')
    standalones = ['X', 'Y', 'Z', 'A', 'B', 'C']
    return name.endswith(suffixes) or name in standalones

# --- NAVIGATION ---
page = st.sidebar.radio("Navigation Menu", ["🏠 Home", "📝 IR Converter", "⚠️ Discrepancy Feature"])

# --- PAGE 1: HOME ---
if page == "🏠 Home":
    st.title("🏠 CMM Quality Suite")
    st.info("Select 'IR Converter' to fill templates or 'Discrepancy Feature' for failure reports.")

# --- PAGE 2: IR CONVERTER ---
elif page == "📝 IR Converter":
    st.title("📝 IR Template Automator")
    uploaded_cmm = st.file_uploader("Upload CMM Result (Excel)", type=["xlsx"], key="ir1")
    uploaded_template = st.file_uploader("Upload IR Template (Excel)", type=["xlsx"], key="ir2")

    if uploaded_cmm and uploaded_template:
        if st.button("🚀 Process IR Report"):
            try:
                # 1. Load CMM Data
                df_scan = pd.read_excel(uploaded_cmm, header=None, nrows=30)
                h_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_cmm = pd.read_excel(uploaded_cmm, header=h_idx)
                df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]

                # Group by exact ID (keeps 30.1 as 30.1)
                df_cmm['MATCH_ID'] = df_cmm['CHARACTERISTIC'].apply(lambda x: get_clean_id(x) if not is_coordinate(x) else None)
                cmm_clean = df_cmm.dropna(subset=['MATCH_ID']).copy()
                cmm_clean['ACTUAL'] = pd.to_numeric(cmm_clean['ACTUAL'], errors='coerce')
                cmm_final = cmm_clean.groupby('MATCH_ID')['ACTUAL'].agg(['min', 'max']).reset_index()

                # 2. Fill Template
                book = load_workbook(io.BytesIO(uploaded_template.getvalue()))
                sheet = book.active

                success_count = 0
                for row_idx in range(1, sheet.max_row + 1):
                    cell_val = sheet.cell(row=row_idx, column=1).value
                    template_id = get_clean_id(cell_val)
                    
                    if template_id:
                        match = cmm_final[cmm_final['MATCH_ID'] == template_id]
                        if not match.empty:
                            v_min, v_max = match.iloc[0]['min'], match.iloc[0]['max']
                            final_str = f"{v_min:.4f} / {v_max:.4f}" if v_min != v_max else f"{v_min:.4f}"
                            sheet.cell(row=row_idx, column=3).value = final_str
                            success_count += 1

                out_ir = io.BytesIO()
                book.save(out_ir)
                st.success(f"✅ Matched {success_count} characteristics!")
                st.download_button("📥 Download Filled IR", out_ir.getvalue(), "Filled_IR.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")

# --- PAGE 3: DISCREPANCY FEATURE ---
elif page == "⚠️ Discrepancy Feature":
    st.title("⚠️ Out-of-Tolerance Reporter")
    uploaded_oot = st.file_uploader("Upload CMM Result", type=["xlsx"], key="oot1")

    if uploaded_oot:
        if st.button("🔍 Generate Discrepancy Report"):
            try:
                # SN from F8
                df_sn = pd.read_excel(uploaded_oot, header=None, nrows=10, usecols="F")
                sn_val = df_sn.iloc[7, 0]
                
                df_scan = pd.read_excel(uploaded_oot, header=None, nrows=30)
                h_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_data = pd.read_excel(uploaded_oot, header=h_idx)
                df_data.columns = [str(c).strip().upper() for c in df_data.columns]

                oot_results = {"SN": [sn_val]}
                for _, row in df_data.iterrows():
                    char_name = str(row.get("CHARACTERISTIC", "")).strip()
                    if is_coordinate(char_name) or char_name.lower() == "nan" or char_name == "":
                        continue
                    try:
                        act, nom = float(row['ACTUAL']), float(row['NOMINAL'])
                        u_tol, l_tol = float(row['UPPER TOL']), float(row['LOWER TOL'])
                        if act > (nom + u_tol) or act < (nom + l_tol):
                            if l_tol == 0: t_str = f"+ {abs(u_tol)}"
                            elif u_tol == 0: t_str = f"- {abs(l_tol)}"
                            else: t_str = f"+/- {abs(u_tol)}"
                            oot_results[f"Dim#{char_name} ({nom} {t_str})"] = [f"{act:.4f}"]
                    except: continue

                if len(oot_results) > 1:
                    st.dataframe(pd.DataFrame(oot_results))
                    out_oot = io.BytesIO()
                    with pd.ExcelWriter(out_oot, engine='xlsxwriter') as writer:
                        pd.DataFrame(oot_results).to_excel(writer, index=False)
                    st.download_button("📥 Download Discrepancy Excel", out_oot.getvalue(), "Discrepancy_Report.xlsx")
                else:
                    st.success("✅ No discrepancies found!")
            except Exception as e:
                st.error(f"Error: {e}")
