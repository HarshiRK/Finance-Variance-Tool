import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XYZ MIS Tool", page_icon="📈")

st.title("📈 MIS Monthly Variance Analyzer")

# 1. FILE UPLOADER (Must do this first!)
uploaded_file = st.file_uploader("STEP 1: Upload your Tally Trial Balance (CSV)", type="csv")

if not uploaded_file:
    st.warning("Please upload a CSV file to see the configuration options.")
else:
    # 2. SMART LOADING
    df_raw = pd.read_csv(uploaded_file, header=None)
    
    # Find the 'Particulars' header row
    header_row = 0
    for i in range(len(df_raw)):
        if "Particulars" in df_raw.iloc[i].values:
            header_row = i
            break
            
    df = pd.read_csv(uploaded_file, skiprows=header_row).dropna(how='all', axis=1)
    df.columns = [str(c).strip() for c in df.columns]

    # 3. SIDEBAR SELECTION (This will appear now!)
    st.sidebar.header("STEP 2: Configure Columns")
    all_cols = df.columns.tolist()
    
    ledger_col = st.sidebar.selectbox("Ledger Name Column", all_cols, index=0)
    
    # We try to auto-select Feb and Mar if they exist
    month_1 = st.sidebar.selectbox("Select Base Month (e.g., Feb)", all_cols)
    month_2 = st.sidebar.selectbox("Select Comparison Month (e.g., Mar)", all_cols)

    st.sidebar.markdown("---")
    generate_btn = st.sidebar.button("STEP 3: Generate Analysis")

    if generate_btn:
        # 4. CLEANING & MATH
        def clean_val(x):
            if pd.isna(x): return 0.0
            s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
            try: return float(s)
            except: return 0.0

        report_df = df[[ledger_col, month_1, month_2]].copy()
        report_df.columns = ['Particulars', month_1, month_2]
        
        report_df[month_1] = report_df[month_1].apply(clean_val)
        report_df[month_2] = report_df[month_2].apply(clean_val)
        report_df['Variance'] = report_df[month_2] - report_df[month_1]
        report_df['Change_%'] = (report_df['Variance'] / report_df[month_1].replace(0, 1))

        # 5. EXCEL EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            report_df.to_excel(writer, sheet_name='Variance', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Variance']

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

            worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red_fmt})
            worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn_fmt})

        st.success(f"Successfully analyzed {month_1} vs {month_2}!")
        st.download_button(
            label="📥 Download Professional Excel Report",
            data=output.getvalue(),
            file_name=f"Variance_Analysis_{month_2}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
