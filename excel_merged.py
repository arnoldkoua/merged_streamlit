import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Merger", page_icon=":pencil:")

st.title("Excel Merger Multi Functions")

# Display instructions for the user
st.write("Upload your Excel files below. The files will be merged into a single file and displayed below. You can then download the merged file.")

# Allow user to enter key variable for merging
key_variable = st.text_input("Enter the key variable to use for merging (add variables):", "")

# Allow user to enter the sheet name for merging
sheet_name = st.text_input("Enter the name of the sheet to merge (optional):", "")

# Allow user to upload files
uploaded_files = st.file_uploader("Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True)

# Display the uploaded files to the user
if uploaded_files:
    st.write("Uploaded Files")
    for file in uploaded_files:
        st.write(file.name)

# Function to check if a sheet name exists in an Excel file
def sheet_exists(file_path, sheet_name):
    try:
        wb = load_workbook(file_path, read_only=True)
        return sheet_name in wb.sheetnames
    except Exception as e:
        return False

# Check if the specified sheet_name exists in any of the uploaded files
sheet_name_exists = False
if sheet_name:
    for file in uploaded_files:
        if not sheet_exists(file, sheet_name):
            st.warning(f"Sheet '{sheet_name}' does not exist in '{file.name}'")
            sheet_name_exists = True
            break

# Merge the uploaded files into a single dataframe when the user clicks the button
if st.button("Merge"):
    if uploaded_files:
        all_dataframes = []
        variable_names = None # Initialize variable_names
        for file in uploaded_files:
            if key_variable == "":
                if sheet_name == "":
                    # If both key_variable and sheet_name are empty, use the default active sheet
                    df = pd.read_excel(file)
                else:
                    df = pd.read_excel(file, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file)

            if key_variable == "":
                if variable_names is None:
                    variable_names = set(df.columns)
                elif variable_names != set(df.columns):
                    st.warning("The uploaded files have different variables. The merge may result in unexpected data. Please upload files with the same variables.")
                    break
            all_dataframes.append(df)
        else:
            if key_variable == "":
                merged_dataframe = pd.concat(all_dataframes)
            else:
                merged_dataframe = all_dataframes[0]
                for i in range(1, len(all_dataframes)):
                    merged_dataframe = pd.merge(merged_dataframe, all_dataframes[i], on=key_variable, how="outer")

            # Display the merged dataframe to the user
            st.write("Merged Dataframe")
            st.write(merged_dataframe)

            # Allow user to download the merged file
            with st.spinner('Downloading...'):
                merged_dataframe.to_excel("merged_file.xlsx", index=False)
            st.success('File ready to be downloaded!')
            st.download_button(label="Download Merged File", data=open("merged_file.xlsx", 'rb').read(), file_name="merged_file.xlsx", mime="application/vnd.ms-excel")
    else:
        st.warning("Please upload at least one file.")
