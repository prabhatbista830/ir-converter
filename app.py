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
page = st.sidebar.radio("Navigation Menu", ["🏠 Home", "📝 IR Converter"])

# --- PAGE 1: HOME ---
if page == "🏠 Home":
    st.title("🏠 CMM Quality Suite")
    st.write("Welcome! This tool handles your quality reporting automation.")
    st.info("Select a tool from the sidebar to begin.")

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
                # Load CMM Data
                df_scan = pd.read_excel(uploaded_cmm, header=None, nrows=25)
                header_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
                df_cmm = pd.read_excel(uploaded_cmm, header=header_idx)
                df_cmm.columns = [str(c).strip().upper() for c in df_cmm.columns]

                # Filter Coordinates
                df_cmm['BASE_CHAR'] = df_cmm['CHARACTERISTIC'].apply(lambda x: None if is_coordinate(x) else extract_base_number(x))
                cmm_clean = df_cmm.dropna(subset=['BASE_CHAR']).copy()
                cmm_final = cmm_clean.groupby('BASE_CHAR')['ACTUAL'].agg(['min', 'max']).reset_index()

                # Write to Template
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

                # Download
                out_ir = io.BytesIO()
                book.save(out_ir)
                st.success(f"✅ Matched {count} characteristics!")
                st.download_button("📥 Download Filled IR", out_ir.getvalue(), "Filled_IR_Report.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")
