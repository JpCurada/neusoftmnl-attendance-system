import pandas as pd
import streamlit as st
from datetime import datetime
from functions import CleaningUtils, AnalysisUtils, im
import plotly.express as px
import io


st.set_page_config(page_title="Neusoft MNL", 
                   page_icon=im,
                   layout="wide", 
                   initial_sidebar_state="auto", 
                   menu_items=None)

st.image('images/neusoft_banner.png', use_column_width="auto")

st.header('Attendance Data Pipeline', divider='grey')
st.write("This tool simplifies attendance data management by transforming the raw extracted data and loading it into clear, understandable, and downloadable table.")


file_up1, file_up2, file_up3 = st.columns(3)
attendance_file = file_up1.file_uploader(label="Attendance Raw Data:", type="xlsx")
master_list_file = file_up2.file_uploader(label="Master List:", type="xlsx")
schedule_file = file_up3.file_uploader(label="Schedule:", type="xlsx")

# Read the Excel file using pandas
if attendance_file is not None and master_list_file is not None and schedule_file is not None:

    try:
        master_list_df = CleaningUtils.create_master_employee_list(master_list_file)
        sched_df = CleaningUtils.create_schedule_dataframe(schedule_file)
        attendance_df, date_list = CleaningUtils.generate_attendance_dataframe(attendance_file)

        merged_df = CleaningUtils.incorporate_master_data(attendance_df, master_list_df, date_list)
        applied_codes_df = CleaningUtils.update_attendance_codes(merged_df, sched_df)

        # Clean time data in attendance date columns using the helper function
        applied_codes_df.iloc[:, 10:] = applied_codes_df.iloc[:, 10:].apply(CleaningUtils.clean_time, axis=1)

        # Remove whitespaces from attendance date columns
        for col in applied_codes_df.columns[10:]:
            applied_codes_df[col] = applied_codes_df[col].str.replace(r'\s+', '', regex=True)

        cleaned_df = CleaningUtils.merge_final_attendance_codes(CleaningUtils, applied_codes_df, sched_df)
        grouped_df = CleaningUtils.transform_attendance_data(cleaned_df)

        st.subheader(f"Attendance Report from {datetime.strptime(date_list[0], '%Y-%m-%d').strftime('%B %d, %Y')} to {datetime.strptime(date_list[-1], '%Y-%m-%d').strftime('%B %d, %Y')}", divider='grey')

        mis_count, mul_count, absent_count, late_count = AnalysisUtils.metric_count(grouped_df)
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Missed Punch Count", mis_count)
        m2.metric("Multiple Punches Count", mul_count)
        m3.metric("Absent Count", absent_count)
        m4.metric("Late Count", late_count)


        code = st.radio('Code Status', options=['(MIS)', '(MUL)', '(ABSENT)', '(L)'], horizontal=True)

        manager_df = AnalysisUtils.count_codes_per_manager(grouped_df)
        date_df = AnalysisUtils.count_code_per_date(grouped_df)

        code_per_manager_fig = AnalysisUtils.plot_leaders_by_code_occurrence(manager_df, code)
        code_per_date_fig = AnalysisUtils.plot_code_occurrence_by_date(date_df, code)

        viz1, viz2 = st.columns(2)
        viz1.plotly_chart(code_per_manager_fig, use_container_width=True)
        viz2.plotly_chart(code_per_date_fig, use_container_width=True)

        with st.sidebar:
            st.header("Data Filter", divider='grey')
            st.caption('Select or input options according to your preferences')
            employees = st.multiselect('Employee Name:', cleaned_df['Employee Name'].unique(), placeholder='')
            lob = st.multiselect('LOB:', cleaned_df['LOB'].unique(), placeholder='')
            shift = st.multiselect('Shift:', cleaned_df['Shift'].unique(), placeholder='')
            site = st.multiselect('Site:', cleaned_df['Site'].unique(), placeholder='')
            leader = st.multiselect('Leader:', cleaned_df['Manager'].unique(), placeholder='')
            employer = st.multiselect('Employer:', cleaned_df['Employer'].unique(), placeholder='')

            st.markdown('---')
            st.caption('@Neusoft')

        if employees:
            grouped_df = grouped_df[grouped_df['Employee Name'].isin(employees)]
        if lob:
            grouped_df = grouped_df[grouped_df['LOB'].isin(lob)]
        if shift:
            grouped_df = grouped_df[grouped_df['Shift'].isin(shift)]
        if site:
            grouped_df = grouped_df[grouped_df['Site'].isin(site)]
        if leader:
            grouped_df = grouped_df[grouped_df['Manager'].isin(leader)]
        if employer:
            grouped_df = grouped_df[grouped_df['Employer'].isin(employer)]    


        with st.spinner('Processing'):

            remarks_count_df = grouped_df['Remarks'].value_counts().rename_axis('Remarks').reset_index(name='Count')
            max_count_value = remarks_count_df['Count'].max()
            figure = px.bar(
                remarks_count_df.sort_values(by='Count', ascending=True),
                x="Count",
                y="Remarks",
                orientation='h',
                title="Remarks' Value Counts"
            )
            st.plotly_chart(figure, use_container_width=True)

            st.subheader("Cleaned Data", divider='grey')
            st.write('Click the arrow at the upper-left corner to view the Filter pane of this data.')

            # buffer to use for excel writer
            buffer = io.BytesIO()

            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Write the styled DataFrame to Excel
                grouped_df.to_excel(writer, sheet_name='Neusoft_MNL_attendance)', index=False)

            # Reset the buffer position
            buffer.seek(0)

            # Displaying the new DataFrame with styled cells
            st.dataframe(grouped_df, use_container_width=True)

            # Provide a download button
            download = st.download_button(
                label="Download Data as Excel",
                data=buffer,
                file_name='neusoft_mnl_attendance.xlsx',
                mime='application/vnd.ms-excel'
            )
    except ValueError:
        # Handle ValueError
        st.error('You uploaded mismatched files. Make sure to upload files to their corresponding File Uploader tab.', icon="🚨")
    except KeyError:
        # Handle KeyError
        st.error('Please check the dates within Raw Attendance Data and Schedule Data. They must be compatible or within each other.', icon="🚨")
