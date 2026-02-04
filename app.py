import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="OOT Checker + Guard Test", layout="wide")

st.title("⚠️ Discrepancy Feature (With Coordinate Guard)")
st.write("Checking SN from **F8** and ignoring **X, Y, Z** coordinates.")

uploaded_cmm = st.file_uploader("Upload CMM Result (Excel)", type=["xlsx"])

if uploaded_cmm:
    try:
        # 1. SN EXTRACTION FROM F8
        df_sn = pd.read_excel(uploaded_cmm, header=None, nrows=10, usecols="F")
        sn_value = df_sn.iloc[7, 0] 
        st.info(f"📍 **Detected SN:** {sn_value}")

        # 2. FIND THE DATA HEADER
        df_scan = pd.read_excel(uploaded_cmm, header=None, nrows=25)
        header_idx = next((i for i, row in df_scan.iterrows() if "CHARACTERISTIC" in row.astype(str).str.upper().values), 12)
        
        df = pd.read_excel(uploaded_cmm, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]

        # 3. RUN THE MATH WITH THE GUARD
        oot_results = {"SN": [sn_value]}
        
        # Suffixes to ignore
        ignore_list = ('.X', '.Y', '.Z', '.A', '.B', '.C', ' X', ' Y', ' Z')

        for _, row in df.iterrows():
            try:
                char_name = str(row.get("CHARACTERISTIC", ""))
                
                # --- THE COORDINATE GUARD ---
                # Skip if it ends with any coordinate suffix
                if char_name.upper().endswith(ignore_list):
                    continue
                
                actual = float(row.get("ACTUAL", 0))
                nominal = float(row.get("NOMINAL", 0))
                u_tol = float(row.get("UPPER TOL", 0))
                l_tol = float(row.get("LOWER TOL", 0))

                # Math check: Out of Tolerance?
                if actual > (nominal + u_tol) or actual < (nominal + l_tol):
                    header_text = f"Dim#{char_name} ({nominal} +/- {abs(u_tol)})"
                    oot_results[header_text] = [actual]
            except:
                continue 

        # 4. SHOW RESULTS
        if len(oot_results) > 1:
            st.success(f"🔥 {len(oot_results)-1} Discrepancies Found (Coordinates Filtered Out)!")
            final_df = pd.DataFrame(oot_results)
            st.dataframe(final_df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False)
            output.seek(0)
            
            st.download_button("📥 Download OOT Excel", output, "OOT_Report.xlsx")
        else:
            st.warning("No OOT values found (or they were all filtered-out coordinates).")

    except Exception as e:
        st.error(f"Something went wrong: {e}")
