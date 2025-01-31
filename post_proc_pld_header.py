import streamlit as st
import pandas as pd
import io

def duplicate_rows_and_modify_specific_sheet(uploaded_file):
    """
    Duplicates existing rows in the 'Rules-Header' sheet of the uploaded Excel file,
    modifies 'Ruleset ShortName', 'Action', and 'Variant Type', and maintains sheet order.

    Args:
        uploaded_file: The uploaded Excel file from Streamlit.

    Returns:
        Tuple (modified_file, original_filename) as BytesIO for download.
    """

    # Load the Excel file
    excel_file = pd.ExcelFile(uploaded_file)

    # Define the target sheet and its desired position
    target_sheet = "Rules-Header"
    before_sheet = "PCRF"
    after_sheet = "Rules-Alias"

    if target_sheet not in excel_file.sheet_names:
        st.error(f"Sheet '{target_sheet}' not found in {uploaded_file.name}")
        return None, None

    # Read the target sheet
    df = excel_file.parse(target_sheet)

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

    # Save to BytesIO instead of a file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet_order = []
        
        # Write sheets in the correct order
        for sheet in excel_file.sheet_names:
            if sheet == after_sheet:
                sheet_order.append(sheet)
                if target_sheet not in sheet_order:
                    modified_df.to_excel(writer, sheet_name=target_sheet, index=False)
                    sheet_order.append(target_sheet)
            elif sheet == before_sheet and target_sheet not in sheet_order:
                modified_df.to_excel(writer, sheet_name=target_sheet, index=False)
                sheet_order.append(target_sheet)
                sheet_order.append(sheet)
            elif sheet != target_sheet:
                excel_file.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)
                sheet_order.append(sheet)
        
        # If "Rules-Header" wasn't placed, add it at the end as a fallback
        if target_sheet not in sheet_order:
            modified_df.to_excel(writer, sheet_name=target_sheet, index=False)

    output.seek(0)
    return output, uploaded_file.name  # Return modified file and original name

# Streamlit UI
st.title("Bulk Excel Sheet Modifier - Maintain Sheet Order")

# File Upload (Multiple Files Allowed)
uploaded_files = st.file_uploader("Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        modified_file, original_filename = duplicate_rows_and_modify_specific_sheet(uploaded_file)

        if modified_file:
            # Provide a download button for each file
            st.download_button(
                label=f"Download Modified: {original_filename}",
                data=modified_file,
                file_name=original_filename,  # Keep the same filename
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
