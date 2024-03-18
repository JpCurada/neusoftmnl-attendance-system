import pandas as pd
import streamlit as st
from functions import prep_attendance_data, merge_master_list, process_schedule_data, apply_codes, clean_time, add_final_codes, apply_color, count_logs, process_masterlist, im
import io

st.set_page_config(page_title="Neusoft MNL", 
                   page_icon=im,
                   layout="wide", 
                   initial_sidebar_state="auto", 
                   menu_items=None)

st.image('images/neusoft_banner.png', use_column_width="auto")

st.header('Attendance Pipeline', divider='grey')
st.write("This tool simplifies attendance data management by transforming the raw extracted data and loading it into clear, understandable, and downloadable table.")

file_up1, file_up2, file_up3 = st.columns(3)
excel_file = file_up1.file_uploader(label="Attendance Raw Data:", type="xlsx")
raw_master_list = file_up2.file_uploader(label="Master List:", type="xlsx")
schedule_file = file_up3.file_uploader(label="Schedule:", type="xlsx")


# Read the Excel file using pandas
if excel_file is not None and raw_master_list is not None and schedule_file is not None:

    file_path = pd.ExcelFile(excel_file)
    attendance_raw_df = pd.read_excel(file_path, sheet_name="打卡时间", engine='openpyxl')
    master_list = process_masterlist(raw_master_list)
    sched_df = process_schedule_data(schedule_file )

    

    prep_df = prep_attendance_data(attendance_raw_df)
    merged_df = merge_master_list(prep_df, master_list)
    applied_codes_df = apply_codes(merged_df, sched_df)

    applied_codes_df.iloc[:, 9:] = applied_codes_df.iloc[:, 9:].apply(clean_time, axis=1)

    attendance_date_cols = applied_codes_df.columns[9:]
    for col in attendance_date_cols:
        applied_codes_df[col] = applied_codes_df[col].str.replace(r'\s+', '', regex=True)


    cleaned_df = add_final_codes(applied_codes_df , sched_df)


    with st.sidebar:
        st.header("Data Filter", divider='grey')
        st.caption('Select or input options according to your preferences')
        employees = st.multiselect('Employee Name:', cleaned_df['Employee Name'].unique(), placeholder='')
        lob = st.multiselect('LOB:', cleaned_df['LOB'].unique(), placeholder='')
        shift = st.multiselect('Shift:', cleaned_df['Shift'].unique(), placeholder='')
        site = st.multiselect('Site:', cleaned_df['Site'].unique(), placeholder='')
        leader = st.multiselect('Leader:', cleaned_df['Leader'].unique(), placeholder='')
        employer = st.multiselect('Employer:', cleaned_df['Employer'].unique(), placeholder='')

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
    if leader:
        cleaned_df = cleaned_df[cleaned_df['Leader'].isin(leader)]
    if employer:
        cleaned_df = cleaned_df[cleaned_df['Employer'].isin(employer)]    

    st.subheader(f'Multiple Logs and Missed Punch', divider='grey')

    # Example usage:
    multiple_logs_df, missed_punch_df = count_logs(cleaned_df)

    col1, col2 = st.columns(2)

    col1.data_editor(
    multiple_logs_df,
    column_config={
        "Multiple Logs Count": st.column_config.ProgressColumn(
            "Multiple Logs Count",
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

    # duplicated_rows = find_duplicates_by_wb_work_number(cleaned_df)
    # if duplicated_rows.shape[0] != 0:
    #     with st.expander('CHECK OUT: Raw Data Issue'):
    #         st.write('The Pipeline detected `WB Work Number` duplicates in the uploaded raw data')
    #         st.dataframe(duplicated_rows)

    # st.subheader(f"Attendance Data from {datetime.strptime(cleaned_df.columns[4], '%Y-%m-%d').strftime('%B %d, %Y')} to {datetime.strptime(cleaned_df.columns[-3], '%Y-%m-%d').strftime('%B %d, %Y')}", divider='grey')

    with st.spinner('Processing'):
        date_columns = list(cleaned_df.columns)[9:]
        styled_df = cleaned_df.style.applymap(apply_color, subset=date_columns)
        
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
