import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import re
import io

# --- UNIVERSAL SETTINGS ---
st.set_page_config(page_title="CMM Quality Suite", layout="wide")

# --- HELPER FUNCTIONS ---
def extract_base_number(text):
    if pd.isna(text): return None
    match = re.search(r'(\d+)', str(text))
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
    st.info("Select a tool from the sidebar to begin.")

# --- PAGE 2: IR CONVERTER (YOUR WORKING LOGIC) ---
elif page == "📝 IR Converter":
    st.title("📝 IR Template Automator")
    uploaded_cmm = st.file_uploader("Upload CMM Result (Excel)", type=["xlsx"], key="ir_cmm")
    uploaded_template = st.file_uploader("Upload IR Template (Excel)", type=["xlsx"], key="ir_tmp")

    if uploaded_cmm and uploaded_template:
        if st.button("🚀 Process IR Report"):
            try:
                # 1. Load CMM Data
                df_scan = pd.read_excel(uploaded_cmm, header=None, nrows=25)
                header_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_cmm = pd.read_excel(uploaded_cmm, header=header_idx)
                df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]

                # 2. Filter Coordinates and Group
                df_cmm['BASE_CHAR'] = df_cmm['CHARACTERISTIC'].apply(lambda x: None if is_coordinate(x) else extract_base_number(x))
                cmm_clean = df_cmm.dropna(subset=['BASE_CHAR']).copy()
                cmm_final = cmm_clean.groupby('BASE_CHAR')['ACTUAL'].agg(['min', 'max']).reset_index()

                # 3. Write to Template
                book = load_workbook(io.BytesIO(uploaded_template.getvalue()))
                sheet = book.active

                count = 0
                for row_idx in range(1, sheet.max_row + 1):
                    cell_value = sheet.cell(row=row_idx, column=1).value
                    base_num = extract_base_number(cell_value)
                    
                    if base_num:
                        match = cmm_final[cmm_final['BASE_CHAR'] == base_num]
                        if not match.empty:
                            v_min, v_max = match.iloc[0]['min'], match.iloc[0]['max']
                            output_str = f"{v_min:.4f} / {v_max:.4f}" if v_min != v_max else f"{v_min:.4f}"
                            sheet.cell(row=row_idx, column=3).value = output_str
                            count += 1

                out_ir = io.BytesIO()
                book.save(out_ir)
                st.success(f"✅ Matched {count} characteristics!")
                st.download_button("📥 Download Filled IR", out_ir.getvalue(), "Filled_IR_Report.xlsx")
            except Exception as e:
                st.error(f"IR Error: {e}")

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
                
                # 2. Process Data
                df_scan = pd.read_excel(uploaded_oot, header=None, nrows=25)
                h_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_data = pd.read_excel(uploaded_oot, header=h_idx)
                df_data.columns = [str(c).strip().upper() for c in df_data.columns]

                oot_results = {"SN": [sn_val]}
                for _, row in df_data.iterrows():
                    char_name = str(row.get("CHARACTERISTIC", ""))
                    if is_coordinate(char_name) or char_name == "nan":
                        continue
                    try:
                        act, nom = float(row['ACTUAL']), float(row['NOMINAL'])
                        u_tol, l_tol = float(row['UPPER TOL']), float(row['LOWER TOL'])
                        if act > (nom + u_tol) or act < (nom + l_tol):
                            # Correct header display logic
                            if l_tol == 0: t_str = f"+ {abs(u_tol)}"
                            elif u_tol == 0: t_str = f"- {abs(l_tol)}"
                            else: t_str = f"+/- {abs(u_tol)}"
                            
                            col_header = f"Dim#{char_name} ({nom} {t_str})"
                            oot_results[col_header] = [f"{act:.4f}"]
                    except: continue

                if len(oot_results) > 1:
                    oot_df = pd.DataFrame(oot_results)
                    st.dataframe(oot_df)
                    out_oot = io.BytesIO()
                    with pd.ExcelWriter(out_oot, engine='xlsxwriter') as writer:
                        oot_df.to_excel(writer, index=False)
                    st.download_button("📥 Download Discrepancy Excel", out_oot.getvalue(), "Discrepancy_Report.xlsx")
                else:
                    st.success("✅ No discrepancies found!")
            except Exception as e:
                st.error(f"Discrepancy Error: {e}")
