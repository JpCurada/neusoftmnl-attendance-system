import pandas as pd
import streamlit as st
from functions import process_attendance_data, apply_combined_styles, merge_master_list, count_logs, find_duplicates_by_wb_work_number, im, master_list
import io
from datetime import datetime

st.set_page_config(page_title="Neusoft MNL", 
                   page_icon=im,
                   layout="wide", 
                   initial_sidebar_state="auto", 
                   menu_items=None)

st.image('images/neusoft_banner.png', use_column_width="auto")

st.header('Attendance ETL Pipeline', divider='grey')
st.write("This tool simplifies attendance data management by transforming the raw extracted data and loading it into clear, understandable, and downloadable table.")

excel_file = st.file_uploader(label="Choose a file:", type="xlsx")

# Read the Excel file using pandas
if excel_file is not None:
    file_path = pd.ExcelFile(excel_file)
    df = pd.read_excel(file_path, sheet_name="打卡时间", engine='openpyxl')

    # Cleaning, processing, and creating a new DataFrame
    with_wb_num_df = process_attendance_data(df)  
    cleaned_df = merge_master_list(with_wb_num_df, master_list)

    for col in cleaned_df.columns:
        cleaned_df[col] = cleaned_df[col].str.replace('外勤','')

    with st.sidebar:
        st.header("Data Filter", divider='grey')
        st.caption('Select or input options according to your preferences')
        employees = st.multiselect('Employee Name:', cleaned_df['Employee Name'].unique(), placeholder='')
        lob = st.multiselect('LOB:', cleaned_df['LOB'].unique(), placeholder='')
        shift = st.multiselect('Shift:', cleaned_df['Shift'].unique(), placeholder='')
        site = st.multiselect('Site:', cleaned_df['Site'].unique(), placeholder='')

        st.markdown('---')
        st.caption('@Neusoft')

    if employees:
        cleaned_df = cleaned_df[cleaned_df['Employee Name'].isin(employees)]
    if lob:
        cleaned_df = cleaned_df[cleaned_df['LOB'].isin(lob)]
    if shift:
        cleaned_df = cleaned_df[cleaned_df['Shift'].isin(shift)]
    if site:
        cleaned_df = cleaned_df[cleaned_df['Site'].isin(site)]
    
    st.subheader(f'No Log and Missed Punch', divider='grey')

    # Example usage:
    no_log_df, missed_punch_df = count_logs(cleaned_df)
  

    col1, col2 = st.columns(2)

    col1.data_editor(
    no_log_df,
    column_config={
        "No Log Count": st.column_config.ProgressColumn(
            "No Log Count",
            format="%f Days",
            min_value=0,
            max_value=30,
        ),
        },
        hide_index=True, 
        use_container_width=True
    )

    col2.data_editor(
    missed_punch_df,
    column_config={
        "Missed Punch Count": st.column_config.ProgressColumn(
            "Missed Punch Count",
            format="%f Days",
            min_value=0,
            max_value=30,), 
            },
        hide_index=True, 
        use_container_width=True
    )

    # buffer to use for excel writer
    n_buffer = io.BytesIO()

    with pd.ExcelWriter(n_buffer, engine='xlsxwriter') as writer:
        # Write the styled DataFrame to Excel
        no_log_df.to_excel(writer, sheet_name='no_log', index=False)
        missed_punch_df.to_excel(writer, sheet_name='missed_punch', index=False)

    # Reset the buffer position
    n_buffer.seek(0)

    # Provide a download button
    download = st.download_button(
        label="Download Data as Excel",
        data=n_buffer,
        file_name='neusoft_mnl_attendance_no_log_missed_punch.xlsx',
        mime='application/vnd.ms-excel'
    )
  

    # buffer to use for excel writer
    buffer = io.BytesIO()

    duplicated_rows = find_duplicates_by_wb_work_number(with_wb_num_df)
    if duplicated_rows.shape[0] != 0:
        with st.expander('CHECK OUT: Raw Data Issue'):
            st.write('The Pipeline detected `WB Work Number` duplicates in the uploaded raw data')
            st.dataframe(duplicated_rows)

    st.subheader(f"Attendance Data from {datetime.strptime(cleaned_df.columns[4], '%Y-%m-%d').strftime('%B %d, %Y')} to {datetime.strptime(cleaned_df.columns[-3], '%Y-%m-%d').strftime('%B %d, %Y')}", divider='grey')

    with st.spinner('Processing'):
        styled_df = cleaned_df.style.map(apply_combined_styles)
        
        # buffer to use for excel writer
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Write the styled DataFrame to Excel
            styled_df.to_excel(writer, sheet_name='Neusoft_MNL_attendance)', index=False)

        # Reset the buffer position
        buffer.seek(0)

        # Displaying the new DataFrame with styled cells
        st.dataframe(styled_df)

        # Provide a download button
        download = st.download_button(
            label="Download Data as Excel",
            data=buffer,
            file_name='neusoft_mnl_attendance.xlsx',
            mime='application/vnd.ms-excel'
        )
