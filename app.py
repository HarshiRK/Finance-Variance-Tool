import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="MIS Automation Portal", page_icon="📊", layout="wide")

st.title("📊 Universal MIS Variance Analyzer")
st.info("Upload your master Trial Balance. The tool will detect all months automatically.")

# 1. UNIVERSAL UPLOADER
uploaded_file = st.file_uploader("Upload Master File (Excel or CSV)", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # 2. READ FILE (CSV or Excel)
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None).fillna("")
        else:
            df_raw = pd.read_excel(uploaded_file, header=None).fillna("")

        # 3. FIND DATA HEADERS
        header_row = None
        for i in range(len(df_raw)):
            if any("Particulars" in str(val) for val in df_raw.iloc[i].values):
                header_row = i
                break

        if header_row is None:
            st.error("Could not find 'Particulars' column. Check the file format.")
            st.stop()

        # 4. MAP MONTHS DYNAMICALLY
        # Tally month labels are usually 2 rows above 'Particulars'
        months_row = df_raw.iloc[max(0, header_row-2)].tolist()
        sub_headers = df_raw.iloc[header_row].tolist()
        
        current_month = ""
        combined_columns = []
        for m, s in zip(months_row, sub_headers):
            m_str = str(m).strip()
            s_str = str(s).strip()
            if m_str and m_str.lower() != "nan":
                current_month = m_str
            
            if current_month and s_str and s_str != "Particulars":
                combined_columns.append(f"{current_month} - {s_str}")
            else:
                combined_columns.append(s_str if s_str else f"Col_{len(combined_columns)}")

        # 5. DATA PREPARATION
        df_main = df_raw.iloc[header_row + 1:].copy()
        df_main.columns = combined_columns
        df_main = df_main.loc[:, ~df_main.columns.str.contains('Col_|^0$')]
        
        # Identify columns
        p_col = [c for c in df_main.columns if 'Particulars' in str(c)][0]
        bal_cols = [c for c in df_main.columns if 'Balance' in str(c)]

        # 6. DYNAMIC DROPDOWN SELECTION
        st.sidebar.header("Comparison Settings")
        st.sidebar.write("The tool found these months in your file:")
        
        m1 = st.sidebar.selectbox("Select Base Month (Older)", bal_cols, index=0)
        m2 = st.sidebar.selectbox("Select Comparison Month (Newer)", bal_cols, index=len(bal_cols)-1)

        def clean_currency(x):
            try:
                s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                return float(s)
            except: return 0.0

        # 7. GENERATE ANALYSIS
        report = df_main[[p_col, m1, m2]].copy()
        report.columns = ['Particulars', m1, m2]
        report[m1] = report[m1].apply(clean_currency)
        report[m2] = report[m2].apply(clean_currency)
        report['Variance'] = report[m2] - report[m1]
        report['% Change'] = (report['Variance'] / report[m1].replace(0, 1))

        # 8. PREVIEW AND DOWNLOAD
        st.subheader(f"Analysis: {m1} vs {m2}")
        st.dataframe(report, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            report.to_excel(writer, sheet_name='MIS_Variance', index=False)
            wb, ws = writer.book, writer.sheets['MIS_Variance']
            
            # Professional Formatting
            hdr = wb.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
            num = wb.add_format({'num_format': '#,##0.00'})
            pct = wb.add_format({'num_format': '0.0%'})
            red = wb.add_format({'bg_color': '#F4CCCC', 'font_color': '#990000'})
            grn = wb.add_format({'bg_color': '#D9EAD3', 'font_color': '#38761D'})

            ws.set_column('A:A', 45)
            ws.set_column('B:D', 18, num)
            ws.set_column('E:E', 12, pct)
            for i, col in enumerate(report.columns):
                ws.write(0, i, col, hdr)
            
            ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red})
            ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn})

        st.download_button(f"📥 Download {m2} Variance Report", output.getvalue(), "MIS_Report.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
