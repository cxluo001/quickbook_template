import pandas as pd
import streamlit as st
from io import BytesIO

def transform_excel_to_csv(input_file):
    # Load the Excel file
    xls = pd.ExcelFile(input_file)
    sheet_name = xls.sheet_names[0]  # Assuming the first sheet is to be processed
    input_df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Define column mappings based on instructions
    supplier_col = "直营\n(Driver)"
    amount_cols = [
        "国际单量收入\n(International Income)",
        "本地单量收入\n(Domestic Income w/o Tax)",
        "总补贴\n(Total Allowance)",
        "总扣款\n(Total Penalty Fine)",
    ]
    memo_value = "Contractor drivers T4A salary"
    account_value = "Operating cost:Cost of Labour-Contractor driver"
    line_descriptions = [
        "T4A contractor drivers,  International packages",
        "T4A contractor drivers,  Domestic packages",
        "T4A contractor drives allowance",
        "T4A contractor drives Deductible",
    ]
    tax_code_value = "Zero-rated"
    hst_tax_code = "HST ON"
    gst_tax_code = "GST"

    # Filter out rows where column A is blank or contains 'Total' or column B is 0
    input_df = input_df[(input_df[supplier_col].notna() & ~input_df[supplier_col].astype(str).str.contains("Total", case=False)) & (input_df.iloc[:, 1] != 0)]

    # Prepare the transformed data
    output_rows = []
    for _, row in input_df.iterrows():
        supplier_value = f"Driver {str(row[supplier_col])}"  # Ensure supplier is treated as a string
        
        # Extract numbers from supplier column and check if any is greater than 10000
        numbers = [float(num) for num in str(row[supplier_col]).split() if num.replace('.', '', 1).isdigit()]
        contains_large_number = any(num > 10000 for num in numbers)
        
        for i, amount_col in enumerate(amount_cols):
            line_amount = row[amount_col] if pd.notna(row[amount_col]) else 0  # Replace NaN with 0
            memo = memo_value if i == 0 else ""  # Memo only for the first entry
            tax_code = tax_code_value
            
            # Set second row relative to vendor to 'HST ON' unless condition for GST is met
            if i == 1:
                tax_code = gst_tax_code if contains_large_number else hst_tax_code

            output_rows.append([
                None,  # *BillNo (empty)
                supplier_value,  # *Supplier
                None,  # *BillDate (empty)
                None,  # *DueDate (empty)
                None,  # Terms (empty)
                None,  # Location (empty)
                memo,  # Memo (first row only)
                account_value,  # *Account
                line_descriptions[i],  # LineDescription
                line_amount,  # *LineAmount
                tax_code,  # *LineTaxCode
                None  # LineTaxAmount (empty)
            ])

    # Convert to DataFrame with correct column names
    output_df = pd.DataFrame(output_rows, columns=[
        "*BillNo", "*Supplier", "*BillDate", "*DueDate", "Terms", "Location",
        "Memo", "*Account", "LineDescription", "*LineAmount", "*LineTaxCode", "LineTaxAmount"
    ])
    
    # Ensure *Supplier column is treated as string
    output_df["*Supplier"] = output_df["*Supplier"].astype(str)
    
    # Save as CSV in memory
    output_buffer = BytesIO()
    output_df.to_csv(output_buffer, index=False)
    output_buffer.seek(0)
    return output_buffer

# Streamlit App
st.title("❤️来自罗爸爸的爱❤️\n司机工资总表模版转换器")
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    st.success("File uploaded successfully!")
    output_csv = transform_excel_to_csv(uploaded_file)
    st.download_button(
        label="Download Transformed CSV",
        data=output_csv,
        file_name="sample_bill_import_tax.csv",
        mime="text/csv"
    )
