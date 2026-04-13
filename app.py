import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Universal MIS Variance Tool", page_icon="📈")

st.title("📈 Universal Monthly Variance Analyzer")
st.markdown("""
This tool works with any Tally-exported Trial Balance. 
1. **Upload** your file. 
2. **Select** the columns you want to compare.
3. **Download** the beautified Excel.
""")

uploaded_file = st.file_uploader("Upload Trial Balance (CSV)", type="csv")

if uploaded_file:
    # 1. SMART LOADING: Find where the data actually starts
    df_raw = pd.read_csv(uploaded_file, header=None)
    
    # We look for the row that has 'Particulars' in it
    header_row = 0
    for i in range(len(df_raw)):
        if "Particulars" in df_raw.iloc[i].values:
            header_row = i
            break
            
    df = pd.read_csv(uploaded_file, skiprows=header_row).dropna(how='all', axis=1)
    df.columns = [str(c).strip() for c in df.columns]

    # 2. USER SELECTION: Let the user pick columns
    all_cols = df.columns.tolist()
    
    st.sidebar.header("Configure Report")
    ledger_col = st.sidebar.selectbox("Select Ledger Name Column", all_cols, index=0)
    month_1 = st.sidebar.selectbox("Select First Month (e.g. Feb)", all_cols)
    month_2 = st.sidebar.selectbox("Select Second Month (e.g. Mar)", all_cols)

    if st.sidebar.button("Generate Analysis"):
        # 3. CLEANING & MATH
        def clean_val(x):
            if pd.isna(x): return 0.0
            s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
            try: return float(s)
            except: return 0.0

        report_df = df[[ledger_col, month_1, month_2]].copy()
        report_df.columns = ['Particulars', 'Month_1_Bal', 'Month_2_Bal']
        
        report_df['Month_1_Bal'] = report_df['Month_1_Bal'].apply(clean_val)
        report_df['Month_2_Bal'] = report_df['Month_2_Bal'].apply(clean_val)
        report_df['Variance'] = report_df['Month_2_Bal'] - report_df['Month_1_Bal']
        report_df['Change_%'] = (report_df['Variance'] / report_df['Month_1_Bal'].replace(0, 1))

        # 4. EXCEL BEAUTIFICATION
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            report_df.to_excel(writer, sheet_name='Variance', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Variance']

            # Professional Styling
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

            # Red for Increase, Green for Decrease
            worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red_fmt})
            worksheet.conditional_format(1, 3, len(report_df), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn_fmt})

        st.success(f"Analysis complete: Comparing {month_1} vs {month_2}")
        st.download_button(
            label="📥 Download Professional Excel Report",
            data=output.getvalue(),
            file_name=f"Variance_{month_1}_vs_{month_2}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
