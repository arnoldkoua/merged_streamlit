import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Merger", page_icon=":pencil:")

st.title("Excel Merger")

# Display instructions for the user
st.write("Upload your Excel files below. The files will be merged into a single file and displayed below. You can then download the merged file.")

# Allow user to upload files
uploaded_files = st.file_uploader("Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True)

# Merge the uploaded files into a single dataframe
if uploaded_files:
    all_dataframes = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        all_dataframes.append(df)
    merged_dataframe = pd.concat(all_dataframes)

    # Display the merged dataframe to the user
    st.write("Merged Dataframe")
    st.write(merged_dataframe)

    # Allow user to download the merged file
    with st.spinner('Downloading...'):
        merged_dataframe.to_excel("merged_file.xlsx", index=False)
    st.success('Download Completed!')
    st.download_button(label="Download Merged File", data=open("merged_file.xlsx", 'rb').read(), file_name="merged_file.xlsx", mime="application/vnd.ms-excel")
