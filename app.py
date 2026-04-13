import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XYZ MIS Tool", page_icon="📊")

st.title("📊 MIS Monthly Variance Analyzer")

uploaded_file = st.file_uploader("STEP 1: Upload your Tally Trial Balance (CSV)", type="csv")

if not uploaded_file:
    st.info("Please upload your Tally CSV file to begin.")
else:
    try:
        # 1. LOAD RAW DATA
        df_raw = pd.read_csv(uploaded_file, header=None).astype(str)
        
        # 2. FIND HEADER ROW (Searching for 'Particulars')
        header_idx = None
        for i in range(len(df_raw)):
            if any("Particulars" in str(val) for val in df_raw.iloc[i].values):
                header_idx = i
                break
        
        if header_idx is None:
            st.error("Error: Could not find 'Particulars' header in this file.")
            st.stop()

        # 3. BUILD CLEAN COLUMN NAMES
        # Usually row -2 has months, row -1 has sub-headers
        months = df_raw.iloc[header_idx-2].replace('nan', None).ffill().tolist()
        sub_headers = df_raw.iloc[header_idx].tolist()
        
        clean_cols = []
        for m, s in zip(months, sub_headers):
            m_txt = str(m).strip() if m and str(m) != 'nan' else ""
            s_txt = str(s).strip() if s and str(s) != 'nan' else "Value"
            
            if m_txt and s_txt != "Particulars":
                clean_cols.append(f"{m_txt} - {s_txt}")
            else:
                clean_cols.append(s_txt)

        # 4. LOAD ACTUAL DATA
        df = pd.read_csv(uploaded_file, skiprows=header_idx + 1, header=None)
        df.columns = clean_cols
        
        # Remove any empty or garbage columns
        df = df.loc[:, ~df.columns.str.contains('Unnamed|nan|Value')]
        
        # Find the actual Particulars column (it might be named 'Particulars' or 'Particulars.1')
        p_col = [c for c in df.columns if 'Particulars' in c][0]
        all_cols = df.columns.tolist()

        # 5. SIDEBAR CONFIGURATION
        st.sidebar.header("STEP 2: Configure")
        month_options = [c for c in all_cols if 'Balance' in c]
        
        m1 = st.sidebar.selectbox("Base Month (Last)", month_options)
        m2 = st.sidebar.selectbox("Comparison Month (Current)", month_options)

        if st.sidebar.button("STEP 3: Generate Analysis"):
            # 6. CLEANING & MATH
            def to_num(x):
                if pd.isna(x) or str(x) == 'nan': return 0.0
                val = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                try: return float(val)
                except: return 0.0

            # Create final report
            report = df[[p_col, m1, m2]].copy()
            report.columns = ['Particulars', m1, m2]
            
            report[m1] = report[m1].apply(to_num)
            report[m2] = report[m2].apply(to_num)
            report['Variance'] = report[m2] - report[m1]
            report['Change_%'] = (report['Variance'] / report[m1].replace(0, 1))

            # 7. EXCEL FORMATTING
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                report.to_excel(writer, sheet_name='MIS_Variance', index=False)
                wb = writer.book
                ws = writer.sheets['MIS_Variance']

                # Format styles
                hdr = wb.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
                num = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                pct = wb.add_format({'num_format': '0.0%', 'border': 1})
                red = wb.add_format({'bg_color': '#F4CCCC', 'font_color': '#990000'})
                grn = wb.add_format({'bg_color': '#D9EAD3', 'font_color': '#38761D'})

                ws.set_column('A:A', 45)
                ws.set_column('B:D', 18, num)
                ws.set_column('E:E', 12, pct)

                for i, col in enumerate(report.columns):
                    ws.write(0, i, col, hdr)

                # Color coding
                ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red})
                ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn})

            st.success("Analysis Ready!")
            st.download_button("📥 Download Variance Report", output.getvalue(), "MIS_Report.xlsx")

    except Exception as e:
        st.error(f"Error processing file: {e}. Please ensure it is a standard Tally CSV.")
