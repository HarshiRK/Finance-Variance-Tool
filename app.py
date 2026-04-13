import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="XYZ MIS Tool", page_icon="📊")

st.title("📊 Monthly Variance Analyzer")
st.markdown("""
### Instructions:
1. Export your **Monthly Trial Balance** from Tally as a **CSV**.
2. Upload the file below.
3. The tool will automatically compare **February vs March** and highlight changes.
""")

uploaded_file = st.file_uploader("Upload Tally CSV", type="csv")

if uploaded_file:
    try:
        # Load and clean data (skipping the first 6 rows of Tally headers)
        df_raw = pd.read_csv(uploaded_file, header=None, skiprows=6)
        
        # Columns 0 (Name), 24 (Feb), 29 (Mar)
        df = df_raw[[0, 24, 29]].copy()
        df.columns = ['Particulars', 'Feb_Balance', 'Mar_Balance']

        def clean_val(x):
            if pd.isna(x): return 0.0
            s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
            try: return float(s)
            except: return 0.0

        df['Feb_Balance'] = df['Feb_Balance'].apply(clean_val)
        df['Mar_Balance'] = df['Mar_Balance'].apply(clean_val)
        df['Variance'] = df['Mar_Balance'] - df['Feb_Balance']
        df['Change_%'] = (df['Variance'] / df['Feb_Balance'].replace(0, 1))

        # Show a quick preview on the screen
        st.success("File processed successfully!")
        st.dataframe(df.head(10), use_container_width=True)

        # Create the beautified Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Variance Report', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Variance Report']

            # Formatting
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
            num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
            red_fmt = workbook.add_format({'bg_color': '#F4CCCC', 'font_color': '#990000'})
            grn_fmt = workbook.add_format({'bg_color': '#D9EAD3', 'font_color': '#38761D'})

            worksheet.set_column('A:A', 40)
            worksheet.set_column('B:D', 18, num_fmt)
            worksheet.set_column('E:E', 12, pct_fmt)
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_fmt)

            # Red color for Increase, Green for Decrease
            worksheet.conditional_format(1, 3, len(df), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red_fmt})
            worksheet.conditional_format(1, 3, len(df), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn_fmt})

        st.download_button(
            label="📥 Download Professional Variance Report",
            data=output.getvalue(),
            file_name="XYZ_Variance_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error: Please ensure you are uploading the correct Tally CSV format. Details: {e}")
