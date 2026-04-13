import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XYZ MIS Tool", page_icon="📊")

st.title("📊 MIS Monthly Variance Analyzer")

uploaded_file = st.file_uploader("STEP 1: Upload Tally CSV", type="csv")

if uploaded_file:
    try:
        # 1. LOAD RAW DATA
        df_raw = pd.read_csv(uploaded_file, header=None).astype(str)
        
        # 2. FIND THE ROW WITH 'Particulars'
        header_idx = None
        for i in range(len(df_raw)):
            if any("Particulars" in str(val) for val in df_raw.iloc[i].values):
                header_idx = i
                break
        
        if header_idx is None:
            st.error("Could not find 'Particulars' column.")
            st.stop()

        # 3. FIX COLUMN NAMES
        # Usually Month is 2 rows above, Sub-header is on the Particulars row
        months = df_raw.iloc[header_idx-2].replace('nan', None).ffill().tolist()
        sub_headers = df_raw.iloc[header_idx].tolist()
        
        combined_cols = []
        for m, s in zip(months, sub_headers):
            m_txt = str(m).strip() if m and str(m) != 'nan' else ""
            s_txt = str(s).strip() if s and str(s) != 'nan' else ""
            if m_txt and s_txt and s_txt != "Particulars":
                combined_cols.append(f"{m_txt} - {s_txt}")
            else:
                combined_cols.append(s_txt if s_txt else "Unnamed")

        # 4. LOAD DATA
        df = pd.read_csv(uploaded_file, skiprows=header_idx + 1, header=None)
        df.columns = combined_cols
        df = df.loc[:, ~df.columns.str.contains('Unnamed|nan')]
        
        # Identify the Name column
        p_col = [c for c in df.columns if 'Particulars' in c][0]
        balance_cols = [c for c in df.columns if 'Balance' in c]

        # 5. SIDEBAR
        st.sidebar.header("Step 2: Compare Months")
        m1 = st.sidebar.selectbox("Base Month (Last)", balance_cols)
        m2 = st.sidebar.selectbox("Comparison Month (Current)", balance_cols)

        # 6. CALCULATIONS
        def clean_val(x):
            try:
                s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                return float(s)
            except: return 0.0

        report = df[[p_col, m1, m2]].copy()
        report.columns = ['Particulars', m1, m2]
        report[m1] = report[m1].apply(clean_val)
        report[m2] = report[m2].apply(clean_val)
        report['Variance'] = report[m2] - report[m1]
        report['Change %'] = (report['Variance'] / report[m1].replace(0, 1))

        # 7. SHOW PREVIEW IMMEDIATELY
        st.subheader(f"Analyzing: {m1} vs {m2}")
        st.dataframe(report.style.format({m1: "{:,.2f}", m2: "{:,.2f}", 'Variance': "{:,.2f}", 'Change %': "{:.1%}"}), use_container_width=True)

        # 8. EXCEL DOWNLOAD
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            report.to_excel(writer, sheet_name='MIS_Variance', index=False)
            wb, ws = writer.book, writer.sheets['MIS_Variance']
            
            # Formats
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
            
            ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red})
            ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn})

        st.download_button("📥 Download Final MIS Report", output.getvalue(), "MIS_Variance_Report.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload the Tally CSV to see the report.")
