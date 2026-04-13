import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XYZ MIS Tool", page_icon="📈")

st.title("📊 MIS Monthly Variance Analyzer")

uploaded_file = st.file_uploader("STEP 1: Upload your Tally Trial Balance (CSV)", type="csv")

if not uploaded_file:
    st.info("Please upload your Tally CSV file to begin.")
else:
    try:
        # 1. LOAD THE RAW DATA
        df_raw = pd.read_csv(uploaded_file, header=None).astype(str)
        
        # 2. FIND THE HEADER ROW (Search for 'Particulars')
        header_idx = None
        for i in range(len(df_raw)):
            if "Particulars" in df_raw.iloc[i].values:
                header_idx = i
                break
        
        if header_idx is None:
            st.error("Could not find 'Particulars' in the file. Please check your CSV.")
            st.stop()

        # 3. EXTRACT MONTHS AND HEADERS
        # We assume months are in the row directly above or two rows above Particulars
        months_row = df_raw.iloc[header_idx-2].replace('nan', None).ffill().tolist()
        headers_row = df_raw.iloc[header_idx].tolist()
        
        combined_cols = []
        for m, h in zip(months_row, headers_row):
            m_txt = str(m).strip() if m and str(m) != 'nan' else ""
            h_txt = str(h).strip() if h and str(h) != 'nan' else "Unnamed"
            
            if m_txt and h_txt != "Particulars":
                combined_cols.append(f"{m_txt} - {h_txt}")
            else:
                combined_cols.append(h_txt)

        # 4. LOAD FINAL DATAFRAME
        df = pd.read_csv(uploaded_file, skiprows=header_idx + 1, header=None)
        df.columns = combined_cols
        
        # Remove empty columns
        df = df.loc[:, ~df.columns.str.contains('Unnamed')]
        all_cols = df.columns.tolist()

        # 5. SIDEBAR SELECTION
        st.sidebar.header("STEP 2: Configure Columns")
        ledger_col = st.sidebar.selectbox("Ledger Name Column", all_cols, index=0)
        
        # Filter for 'Balance' columns to make the list shorter
        balance_options = [c for c in all_cols if 'Balance' in c]
        
        month_1 = st.sidebar.selectbox("Base Month (Last)", balance_options if balance_options else all_cols)
        month_2 = st.sidebar.selectbox("Comparison Month (Current)", balance_options if balance_options else all_cols)

        if st.sidebar.button("STEP 3: Generate Analysis"):
            # 6. CLEANING & MATH
            def clean_val(x):
                if pd.isna(x) or str(x) == 'nan': return 0.0
                s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                try: return float(s)
                except: return 0.0

            report_df = df[[ledger_col, month_1, month_2]].copy()
            report_df.columns = ['Particulars', month_1, month_2]
            
            report_df[month_1] = report_df[month_1].apply(clean_val)
            report_df[month_2] = report_df[month_2].apply(clean_val)
            report_df['Variance'] = report_df[month_2] - report_df[month_1]
            report_df['Change_%'] = (report_df['Variance'] / report_df[month_1].replace(0, 1))

            # 7. EXCEL EXPORT
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                report_df.to_excel(writer, sheet_name='Variance', index=False)
                workbook  = writer.book
                worksheet = writer.sheets['Variance']

                # Styling
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
                num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
                pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
                red_fmt = workbook.add_format({'bg_color': '#F4CCCC', 'font_color': '#990000'})
                grn_fmt = workbook.add_format({'bg_color': '#D9EAD3', 'font_color': '#38761D'})

                worksheet.set_column('A:A', 45)
                worksheet.set_column('B:D', 18, num_fmt)
                worksheet.set_column('E:E', 12, pct_fmt)
                
                for col_num, value in enumerate(report_df.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)

                worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red_fmt})
                worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn_fmt})

            st.success("Analysis Ready for Submission!")
            st.download_button(
                label="📥 Download Professional Variance Report",
                data=output.getvalue(),
                file_name=f"MIS_Variance_Report_{month_2}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Please check your CSV format. Error details: {e}")
