import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XYZ MIS Tool", page_icon="📈")

st.title("📈 MIS Monthly Variance Analyzer")

# 1. FILE UPLOADER
uploaded_file = st.file_uploader("STEP 1: Upload your Tally Trial Balance (CSV)", type="csv")

if not uploaded_file:
    st.info("Please upload your Tally CSV file to begin.")
else:
    # 2. LOAD DATA
    # Read the whole thing as strings to prevent errors
    df_raw = pd.read_csv(uploaded_file, header=None).astype(str)
    
    # 3. LOCATE THE MONTH NAMES AND HEADERS
    # Usually Row 3 has months, Row 5 has 'Particulars'
    month_row_idx = 3 
    header_row_idx = 5
    
    # Extract the rows
    months = df_raw.iloc[month_row_idx].replace('nan', None).ffill().tolist()
    headers = df_raw.iloc[header_row_idx].tolist()
    
    # Combine them: "Feb-26 - Balance"
    combined_cols = []
    for m, h in zip(months, headers):
        m_clean = str(m).strip() if m else ""
        h_clean = str(h).strip() if h else ""
        if m_clean and h_clean and h_clean != 'nan':
            combined_cols.append(f"{m_clean} - {h_clean}")
        else:
            combined_cols.append(h_clean if h_clean != 'nan' else "Unnamed")

    # Load the actual data part
    df = pd.read_csv(uploaded_file, skiprows=header_row_idx + 1, header=None)
    df.columns = combined_cols
    
    # Remove any columns that are just 'Unnamed'
    df = df.loc[:, ~df.columns.str.contains('Unnamed')]
    all_cols = df.columns.tolist()

    # 4. SIDEBAR SELECTION
    st.sidebar.header("STEP 2: Configure Columns")
    
    ledger_col = st.sidebar.selectbox("Select Ledger Name Column", all_cols, index=0)
    
    # Filter columns to only show 'Balance' options to make it easier
    balance_cols = [c for c in all_cols if 'Balance' in c]
    
    month_1 = st.sidebar.selectbox("Select Base Month", balance_cols if balance_cols else all_cols)
    month_2 = st.sidebar.selectbox("Select Comparison Month", balance_cols if balance_cols else all_cols)

    st.sidebar.markdown("---")
    if st.sidebar.button("STEP 3: Generate Analysis"):
        # 5. CLEANING & MATH
        def clean_val(x):
            if pd.isna(x): return 0.0
            s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
            try: return float(s)
            except: return 0.0

        report_df = df[[ledger_col, month_1, month_2]].copy()
        report_df.columns = ['Particulars', month_1, month_2]
        
        report_df[month_1] = report_df[month_1].apply(clean_val)
        report_df[month_2] = report_df[month_2].apply(clean_val)
        
        # Calculate Variance
        report_df['Variance_Amt'] = report_df[month_2] - report_df[month_1]
        report_df['Change_%'] = (report_df['Variance_Amt'] / report_df[month_1].replace(0, 1))

        # 6. EXCEL EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            report_df.to_excel(writer, sheet_name='Variance', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Variance']

            # Styles
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
            num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
            red_fmt = workbook.add_format({'bg_color': '#F4CCCC', 'font_color': '#990000'})
            grn_fmt = workbook.add_format({'bg_color': '#D9EAD3', 'font_color': '#38761D'})

            worksheet.set_column('A:A', 40)
            worksheet.set_column('B:D', 18, num_fmt)
            worksheet.set_column('E:E', 12, pct_fmt)
            
            for col_num, value in enumerate(report_df.columns.values):
                worksheet.write(0, col_num, value, header_fmt)

            # Highlighting
            worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red_fmt})
            worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn_fmt})

        st.success("Analysis Complete!")
        st.download_button(
            label="📥 Download Professional Excel Report",
            data=output.getvalue(),
            file_name=f"Variance_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
