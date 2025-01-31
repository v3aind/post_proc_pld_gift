import streamlit as st
import pandas as pd
import io

def process_file(uploaded_file):
    sheet_name = "Rules-Header"
    
    # Load the Excel file
    excel_file = pd.ExcelFile(uploaded_file)
    
    # Read the specified sheet
    df = excel_file.parse(sheet_name)
    
    # Duplicate the DataFrame
    duplicated_df = df.copy()
    
    # Modify specific columns if they exist
    if 'Ruleset ShortName' in duplicated_df.columns:
        duplicated_df['Ruleset ShortName'] = ''  
    
    if 'Variant Type' in duplicated_df.columns:
        duplicated_df['Variant Type'] = 'GF'
    
    if 'Action' in duplicated_df.columns:
        duplicated_df['Action'] = 'INSERT'  
    
    # Concatenate original and duplicated data
    modified_df = pd.concat([df, duplicated_df], ignore_index=True)
    
    # Create an output Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        modified_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Copy other sheets from the original file
        for sheet in excel_file.sheet_names:
            if sheet != sheet_name:
                excel_file.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)
    
    output.seek(0)
    return output

# Streamlit UI
st.title("Excel File Processor")
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        processed_file = process_file(uploaded_file)
        st.download_button(
            label=f"Download modified {uploaded_file.name.replace('.xlsx', '_GIFT.xlsx')}",
            data=processed_file,
            file_name=uploaded_file.name.replace(".xlsx", "_GIFT.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
