import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import re
import io

st.set_page_config(page_title="CMM Quality Suite", layout="wide")

# --- RESTORED ORIGINAL IR LOGIC ---
def extract_base_number(text):
    if pd.isna(text): return None
    match = re.search(r'(\d+)', str(text))
    return match.group(1) if match else None

def is_coordinate(char_name):
    name = str(char_name).strip().upper()
    suffixes = ('.X', '.Y', '.Z', '.A', '.B', '.C', ' X', ' Y', ' Z')
    return name.endswith(suffixes) or name in ['X', 'Y', 'Z', 'A', 'B', 'C']

# --- NAVIGATION ---
page = st.sidebar.radio("Navigation Menu", ["🏠 Home", "📝 IR Converter", "⚠️ Discrepancy Feature"])

# --- PAGE 1: HOME ---
if page == "🏠 Home":
    st.title("🏠 CMM Quality Suite")
    st.write("Stable Version: IR Converter (Original) + Discrepancy Feature.")

# --- PAGE 2: IR CONVERTER (REVERTED TO WORKING VERSION) ---
elif page == "📝 IR Converter":
    st.title("📝 IR Template Automator")
    uploaded_cmm = st.file_uploader("Upload CMM Result", type=["xlsx"], key="ir_1")
    uploaded_template = st.file_uploader("Upload IR Template", type=["xlsx"], key="ir_2")

    if uploaded_cmm and uploaded_template:
        if st.button("🚀 Process IR Report"):
            try:
                # 1. Standard CMM Loading
                df_scan = pd.read_excel(uploaded_cmm, header=None, nrows=30)
                h_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_cmm = pd.read_excel(uploaded_cmm, header=h_idx)
                df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]

                # 2. Original Grouping Logic
                df_cmm['BASE_CHAR'] = df_cmm['CHARACTERISTIC'].apply(lambda x: None if is_coordinate(x) else extract_base_number(x))
                cmm_clean = df_cmm.dropna(subset=['BASE_CHAR']).copy()
                cmm_final = cmm_clean.groupby('BASE_CHAR')['ACTUAL'].agg(['min', 'max']).reset_index()

                # 3. Write to Template
                book = load_workbook(io.BytesIO(uploaded_template.getvalue()))
                sheet = book.active

                count = 0
                for row_idx in range(1, sheet.max_row + 1):
                    cell_val = sheet.cell(row=row_idx, column=1).value
                    base_num = extract_base_number(cell_val)
                    if base_num:
                        match = cmm_final[cmm_final['BASE_CHAR'] == base_num]
                        if not match.empty:
                            v_min, v_max = match.iloc[0]['min'], match.iloc[0]['max']
                            res = f"{v_min:.4f} / {v_max:.4f}" if v_min != v_max else f"{v_min:.4f}"
                            sheet.cell(row=row_idx, column=3).value = res
                            count += 1

                output = io.BytesIO()
                book.save(output)
                st.success(f"✅ IR Restored: {count} matches found.")
                st.download_button("📥 Download IR", output.getvalue(), "Filled_IR.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")

# --- PAGE 3: DISCREPANCY FEATURE ---
elif page == "⚠️ Discrepancy Feature":
    st.title("⚠️ Out-of-Tolerance Reporter")
    uploaded_oot = st.file_uploader("Upload CMM Result", type=["xlsx"], key="oot_1")

    if uploaded_oot:
        if st.button("🔍 Generate Discrepancy Report"):
            try:
                # SN from F8
                df_sn = pd.read_excel(uploaded_oot, header=None, nrows=10, usecols="F")
                sn_val = df_sn.iloc[7, 0]
                
                # Data Load
                df_scan = pd.read_excel(uploaded_oot, header=None, nrows=30)
                h_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_data = pd.read_excel(uploaded_oot, header=h_idx)
                df_data.columns = [str(c).strip().upper() for c in df_data.columns]

                oot_results = {"SN": [sn_val]}
                for _, row in df_data.iterrows():
                    name = str(row.get("CHARACTERISTIC", "")).strip()
                    if is_coordinate(name) or name.lower() == "nan" or name == "":
                        continue
                    try:
                        act, nom = float(row['ACTUAL']), float(row['NOMINAL'])
                        u_tol, l_tol = float(row['UPPER TOL']), float(row['LOWER TOL'])
                        if act > (nom + u_tol) or act < (nom + l_tol):
                            # Header styling
                            if l_tol == 0: t_str = f"+ {abs(u_tol)}"
                            elif u_tol == 0: t_str = f"- {abs(l_tol)}"
                            else: t_str = f"+/- {abs(u_tol)}"
                            
                            oot_results[f"Dim#{name} ({nom} {t_str})"] = [f"{act:.4f}"]
                    except: continue

                if len(oot_results) > 1:
                    st.dataframe(pd.DataFrame(oot_results))
                    out_oot = io.BytesIO()
                    with pd.ExcelWriter(out_oot, engine='xlsxwriter') as writer:
                        pd.DataFrame(oot_results).to_excel(writer, index=False)
                    st.download_button("📥 Download Discrepancies", out_oot.getvalue(), "Discrepancy_Report.xlsx")
                else:
                    st.success("✅ No discrepancies found!")
            except Exception as e:
                st.error(f"Error: {e}")
