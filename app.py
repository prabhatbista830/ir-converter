import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="OOT Checker Test", layout="wide")

st.title("⚠️ Discrepancy Feature Test")
st.write("Checking SN from **F8** and matching **Actual** vs **Nominal**.")

uploaded_cmm = st.file_uploader("Upload CMM Result (Excel)", type=["xlsx"])

if uploaded_cmm:
    try:
        # 1. TEST SN EXTRACTION FROM F8
        # We read just cell F8 (Row index 7, Column F is index 5)
        df_sn = pd.read_excel(uploaded_cmm, header=None, nrows=10, usecols="F")
        sn_value = df_sn.iloc[7, 0] 
        
        st.info(f"📍 **Detected SN (from Cell F8):** {sn_value}")

        # 2. FIND THE DATA HEADER
        # We look for the row containing "Characteristic"
        df_scan = pd.read_excel(uploaded_cmm, header=None, nrows=25)
        header_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
        
        # Load the actual data
        df = pd.read_excel(uploaded_cmm, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]

        # 3. RUN THE MATH
        oot_results = {"SN": [sn_value]}
        
        # We look for these exact column names in your Excel
        for _, row in df.iterrows():
            try:
                char_name = str(row.get("CHARACTERISTIC", ""))
                actual = float(row.get("ACTUAL", 0))
                nominal = float(row.get("NOMINAL", 0))
                u_tol = float(row.get("UPPER TOL", 0))
                l_tol = float(row.get("LOWER TOL", 0))

                # Logic: Check if outside the boundaries
                if actual > (nominal + u_tol) or actual < (nominal + l_tol):
                    header_text = f"Dim#{char_name} ({nominal} +/- {abs(u_tol)})"
                    oot_results[header_text] = [actual]
            except:
                continue # Skip rows that aren't numbers (like text or empty rows)

        # 4. SHOW RESULTS
        if len(oot_results) > 1:
            st.success("🔥 Discrepancies Found!")
            final_df = pd.DataFrame(oot_results)
            st.dataframe(final_df)

            # Excel Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False)
            output.seek(0)
            
            st.download_button("📥 Download OOT Excel", output, "OOT_Report.xlsx")
        else:
            st.warning("No OOT values found based on the math.")

    except Exception as e:
        st.error(f"Something went wrong: {e}")
