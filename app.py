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
    # Guarding against X, Y, Z suffixes and standalone letters
    suffixes = ('.X', '.Y', '.Z', '.A', '.B', '.C', ' X', ' Y', ' Z')
    standalones = ['X', 'Y', 'Z', 'A', 'B', 'C']
    return name.endswith(suffixes) or name in standalones

# --- NAVIGATION ---
page = st.sidebar.radio("Navigation Menu", ["🏠 Home", "📝 IR Converter", "⚠️ Discrepancy Feature"])

# --- PAGE 1: HOME ---
if page == "🏠 Home":
    st.title("🏠 CMM Quality Suite")
    st.write("Welcome! This tool handles your quality reporting automation.")
    st.info("""
    **Available Tools:**
    * **IR Converter:** Fills your standard IR Template with CMM results.
    * **Discrepancy Feature:** Generates a summary of all dimensions that failed tolerance.
    """)

# --- PAGE 2: IR CONVERTER ---
elif page == "📝 IR Converter":
    st.title("📝 IR Template Automator")
    st.write("Upload your files to populate the standard Inspection Report.")
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_cmm = st.file_uploader("Upload CMM Result (Excel)", type=["xlsx"], key="ir_cmm")
    with col2:
        uploaded_template = st.file_uploader("Upload IR Template (Excel)", type=["xlsx"], key="ir_tmp")

    if uploaded_cmm and uploaded_template:
        if st.button("🚀 Process IR Report"):
            try:
                df_scan = pd.read_excel(uploaded_cmm, header=None, nrows=25)
                header_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_cmm = pd.read_excel(uploaded_cmm, header=header_idx)
                df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]

                df_cmm['BASE_CHAR'] = df_cmm['CHARACTERISTIC'].apply(lambda x: None if is_coordinate(x) else extract_base_number(x))
                cmm_clean = df_cmm.dropna(subset=['BASE_CHAR']).copy()
                cmm_final = cmm_clean.groupby('BASE_CHAR')['ACTUAL'].agg(['min', 'max']).reset_index()

                template_bytes = uploaded_template.getvalue()
                book = load_workbook(io.BytesIO(template_bytes))
                sheet = book.active

                count = 0
                for row_idx in range(1, sheet.max_row + 1):
                    cell_value = sheet.cell(row=row_idx, column=1).value
                    base_num = extract_base_number(cell_value)
                    
                    if base_num:
                        match = cmm_final[cmm_final['BASE_CHAR'] == base_num]
                        if not match.empty:
                            val_min = match.iloc[0]['min']
                            val_max = match.iloc[0]['max']
                            output_str = f"{val_min:.4f} / {val_max:.4f}" if val_min != val_max else f"{val_min:.4f}"
                            sheet.cell(row=row_idx, column=3).value = output_str
                            count += 1

                out_ir = io.BytesIO()
                book.save(out_ir)
                st.success(f"✅ Matched {count} characteristics!")
                st.download_button("📥 Download Filled IR", out_ir.getvalue(), "Filled_IR_Report.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")

# --- PAGE 3: DISCREPANCY FEATURE ---
elif page == "⚠️ Discrepancy Feature":
    st.title("⚠️ Out-of-Tolerance Reporter")
    st.write("This tool creates a horizontal report of failing dimensions only.")
    
    uploaded_oot = st.file_uploader("Upload CMM Result", type=["xlsx"], key="oot_cmm")

    if uploaded_oot:
        if st.button("🔍 Generate Discrepancy Report"):
            try:
                # 1. SN from F8
                df_sn = pd.read_excel(uploaded_oot, header=None, nrows=10, usecols="F")
                sn_val = df_sn.iloc[7, 0]
                
                # 2. Process Data - Reading up to 5000 rows to ensure we don't miss the end (174)
                df_scan = pd.read_excel(uploaded_oot, header=None, nrows=30)
                h_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_data = pd.read_excel(uploaded_oot, header=h_idx)
                df_data.columns = [str(c).strip().upper() for c in df_data.columns]

                oot_results = {"SN": [sn_val]}
                
                for _, row in df_data.iterrows():
                    # Get character name and clean it
                    char_name = str(row.get("CHARACTERISTIC", "")).strip()
                    
                    # Skip if empty or coordinate
                    if char_name == "" or char_name.lower() == "nan" or is_coordinate(char_name):
                        continue
                    
                    try:
                        # Convert to float for math
                        act = float(row.get("ACTUAL", 0))
                        nom = float(row.get("NOMINAL", 0))
                        u_tol = float(row.get("UPPER TOL", 0))
                        l_tol = float(row.get("LOWER TOL", 0))

                        # OOT Math Check
                        if act > (nom + u_tol) or act < (nom + l_tol):
                            col_header = f"Dim#{char_name} ({nom} +/- {abs(u_tol)})"
                            oot_results[col_header] = [f"{act:.4f}"]
                    except:
                        continue

                if len(oot_results) > 1:
                    oot_df = pd.DataFrame(oot_results)
                    st.write("### Failure Summary:")
                    st.dataframe(oot_df)

                    out_oot = io.BytesIO()
                    with pd.ExcelWriter(out_oot, engine='xlsxwriter') as writer:
                        oot_df.to_excel(writer, index=False)
                    st.download_button("📥 Download Discrepancy Excel", out_oot.getvalue(), "Discrepancy_Report.xlsx")
                else:
                    st.success("✅ No discrepancies found (excluding coordinates)!")
            except Exception as e:
                st.error(f"Error: {e}")
