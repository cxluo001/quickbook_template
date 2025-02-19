import pandas as pd
import streamlit as st
from io import BytesIO
import csv
import re  # Import the regular expression module

def transform_excel_to_csv(input_file):
    # Load the Excel file
    xls = pd.ExcelFile(input_file)
    sheet_name = xls.sheet_names[0]  # Assuming the first sheet is to be processed
    input_df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
    
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
        "T4A contractor drivers, International packages",
        "T4A contractor drivers, Domestic packages",
        "T4A contractor drives Allowance",
        "T4A contractor drives Deductible",
    ]
    tax_code_value = "Zero-rated"
    hst_tax_code = "HST ON"
    gst_tax_code = "GST"

    # Filter out rows where column A is blank, contains 'Total', or column B is 0
    input_df = input_df[input_df[supplier_col].notna() & ~input_df[supplier_col].astype(str).str.contains("Total", case=False)]
    input_df = input_df[input_df.iloc[:, 1] != "0"]  # Assuming column B is the second column

    # Prepare the transformed data
    output_rows = []
    for _, row in input_df.iterrows():
        supplier_value = row[supplier_col].strip()
        # Extract numeric part from supplier_value, in case there are alphabets
        numeric_part = re.findall(r'\d+', supplier_value)  # Extract all digits
        numeric_part = int(numeric_part[0]) if numeric_part else 0  # Get the first numeric part (if available)
        
        # If the numeric part exceeds 10,000, it will use the GST tax code
        contains_large_number = numeric_part > 10000
        
        # Now, create the supplier name transformation
        if numeric_part > 0:
            supplier_value = f"Driver {supplier_value}"  # Keep the format consistent

        for i, amount_col in enumerate(amount_cols):
            line_amount = row[amount_col] if pd.notna(row[amount_col]) else 0  # Replace NaN with 0
            memo = memo_value if i == 0 else ""  # Memo only for the first entry
            tax_code = tax_code_value
            
            # Set second row relative to vendor to 'HST ON' unless condition for GST is met
            if i == 1:
                tax_code = gst_tax_code if contains_large_number else hst_tax_code

            # Convert to float, round to 2 decimals, and handle errors
            try:
                line_amount = round(float(line_amount), 2)
            except ValueError:
                line_amount = 0  # Default to 0 if not a valid number

            # Negate the amount for the 4th line (Deductible) if it's not zero
            if i == 3 and line_amount != 0:
                line_amount = -abs(line_amount)  # Ensure it's negative

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
                line_amount,  # *LineAmount (rounded to 2 decimals)
                tax_code,  # *LineTaxCode
                None  # LineTaxAmount (empty)
            ])

    # Convert to DataFrame with correct column names
    output_df = pd.DataFrame(output_rows, columns=[
        "*BillNo", "*Supplier", "*BillDate", "*DueDate", "Terms", "Location",
        "Memo", "*Account", "LineDescription", "*LineAmount", "*LineTaxCode", "LineTaxAmount"
    ])
    
    # Save as CSV in memory
    output_buffer = BytesIO()
    output_df.to_csv(output_buffer, index=False, quoting=csv.QUOTE_MINIMAL)
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
