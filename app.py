import pandas as pd
import streamlit as st
from functions import create_list_of_dates, clean_process_create_df, apply_combined_styles, im
import io

st.set_page_config(page_title="Neusoft MNL", 
                   icon=im,
                   layout="wide", 
                   initial_sidebar_state="auto", 
                   menu_items=None)

st.title("Neusoft MNL — Attendance Data")

st.markdown('---')

excel_file = st.file_uploader(label="Choose a file:", type="xlsx")

# Read the Excel file using pandas
if excel_file is not None:
    file_path = pd.ExcelFile(excel_file)
    df = pd.read_excel(file_path, sheet_name="打卡时间", engine='openpyxl')

    # Generating list of dates
    generated_dates = create_list_of_dates(df) 

    # Cleaning, processing, and creating a new DataFrame
    cleaned_df = clean_process_create_df(df, generated_dates)  

    employees = st.multiselect('Select Employees to observe:', cleaned_df.columns[2:])

    if employees:
        columns = ['Date', 'Weekday'] + employees
        cleaned_df = cleaned_df[['Date', 'Weekday'] + employees]

    # Displaying the new DataFrame with styled cells
    styled_df = cleaned_df.style.map(apply_combined_styles)
    st.table(styled_df)

    # buffer to use for excel writer
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write the styled DataFrame to Excel
        styled_df.to_excel(writer, sheet_name='Neusoft_MNL_attendance)', index=False)

    # Reset the buffer position
    buffer.seek(0)

    # Provide a download button
    download = st.download_button(
        label="Download Data as Excel",
        data=buffer,
        file_name='neusoft_mnl_attendance.xlsx',
        mime='application/vnd.ms-excel'
    )
