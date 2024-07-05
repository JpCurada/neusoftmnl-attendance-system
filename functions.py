import pandas as pd
import plotly.express as px
import numpy as np
from datetime import datetime, timedelta
from PIL import Image
import re 

import warnings
warnings.simplefilter("ignore")

pd.options.mode.chained_assignment = None  # Set pandas option to suppress chained assignment warning

im = Image.open("images/neusoft_logo.png")

class CleaningUtils:

  @staticmethod
  def create_master_employee_list(filepath):
    """
    This function reads an Excel file containing active and inactive employee data
    and returns a consolidated DataFrame with specific columns.

    Args:
        filepath (str): Path to the Excel file containing employee data.

    Returns:
        pandas.DataFrame: A DataFrame containing a consolidated list of employees
                          with specified columns.
    """
    # Read active employee data
    active_df = pd.read_excel(filepath, sheet_name='Active')

    # Read inactive employee data
    inactive_df = pd.read_excel(filepath, sheet_name='Inactive')

    # Combine active and inactive dataframes
    combined_df = pd.concat([active_df, inactive_df], ignore_index=True)

    # Select desired columns
    employee_df = combined_df[['Employee Name', 'Employee Code (ID)' ,'WB Work Number','RAG', 'Work Location', 'Shift', 'Site', 'LOB', 'Leader','Employer']]

    return employee_df

  @staticmethod
  def create_schedule_dataframe(filepath):
    """
    This function reads an Excel file containing schedule data from multiple sheets
    and returns a consolidated DataFrame with formatted dates as columns.

    Args:
        filepath (str): Path to the Excel file containing schedule data.

    Returns:
        pandas.DataFrame: A DataFrame containing a consolidated schedule with 
                          formatted dates as columns.
    """

    # Read schedule dates from RBC sheet (assuming first row, column offset 6)
    schedule_dates = pd.read_excel(filepath, sheet_name='RBC').iloc[0, 6:].tolist()
    formatted_dates = [date.strftime('%Y-%m-%d') for date in schedule_dates]

    # Define column names including employee information and formatted dates
    columns = ['Index', 'Employee Number', 'LOB', 'EmployeeID', 'Work Number', 'Name']
    columns.extend(formatted_dates)

    # Read data from each sheet (assuming data starts from row 3)
    rbc_df = pd.read_excel(filepath, sheet_name='RBC')[3:]
    hsq_df = pd.read_excel(filepath, sheet_name='HSQ')[3:]
    # idn_df = pd.read_excel(filepath, sheet_name='IDN')[3:]
    isa_df = pd.read_excel(filepath, sheet_name='ISA')[3:]
    # bz_df = pd.read_excel(filepath, sheet_name='BZ')[3:]
    # Concatenate dataframes and set column names
    df = pd.concat([rbc_df, hsq_df, isa_df], ignore_index=True)
    df.columns = columns

    # Remove whitespaces from schedule date columns
    for date_col in columns[6:]:
      df[date_col] = df[date_col].str.replace(r'\s+', '', regex=True)

    return df

  @staticmethod
  def generate_attendance_dataframe(filepath):
    """
    This function reads an attendance data Excel file and returns a DataFrame
    with formatted dates as columns.

    Args:
        filepath (str): Path to the Excel file containing attendance data.

    Returns:
        pandas.DataFrame: A DataFrame containing attendance data with formatted 
                          dates as columns.
    """

    # Read attendance data from the Excel file
    attendance_df = pd.read_excel(filepath, sheet_name='打卡时间')

    # Extract dates from the first column using regular expressions
    date_strings = re.findall(r'\d{4}-\d{2}-\d{2}', attendance_df.columns[0])
    start_date = datetime.strptime(date_strings[0], '%Y-%m-%d')
    end_date = datetime.strptime(date_strings[-1], '%Y-%m-%d')

    # Generate a list of dates between start and end date (inclusive)
    date_range = (end_date - start_date).days + 1
    date_list = [(start_date + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(date_range)]

    # Select and copy data from the third row onwards
    sliced_df = attendance_df[2:].copy()

    # Define column names including employee information and formatted dates
    columns = ["Employee Name", "Attendance Group", "Department", "WB Work Number", "Position", "User ID"]
    columns.extend(date_list)

    # Set column names for the sliced dataframe
    sliced_df.columns = columns

    return sliced_df, date_list
  
  @staticmethod
  def incorporate_master_data(cleaned_df, master_list, date_list):
    """
    Merges employee information from the master list into the cleaned dataframe.

    Args:
        cleaned_df (pd.DataFrame): The DataFrame to be enriched with master data.
        master_list (pd.DataFrame): The DataFrame containing additional employee information.

    Returns:
        pd.DataFrame: The enhanced DataFrame with merged employee data.
    """

    # Initialize lists to store data for new columns
    employee_names = []  # Clearer variable name
    employee_ids = []
    work_numbers = []  # Consistent naming convention
    rags = []
    work_locations = []
    lobs = []  # Abbreviation for Line of Business
    sites = []
    shifts = []
    managers = []  # More descriptive name
    employers = []

    # Iterate through each row in the cleaned dataframe
    for index, row in cleaned_df.iterrows():
        extracted_digits = re.findall(r'\d+', str(row.get('WB Work Number', '')))
        digits = ''.join(extracted_digits)
        work_number = 'WB' + digits  # Store for clarity

        # Check for matching work number in the master list
        matching_record = master_list.loc[master_list['WB Work Number'] == work_number]

        if not matching_record.empty:
            # Extract data from the matching master record
            employee_names.append(matching_record.iloc[0]['Employee Name'])
            employee_ids.append(matching_record.iloc[0]['Employee Code (ID)'])
            work_numbers.append(work_number)  # Reuse stored value
            rags.append(matching_record.iloc[0]['RAG'])
            work_locations.append(matching_record.iloc[0]['Work Location'])
            lobs.append(matching_record.iloc[0]['LOB'])
            sites.append(matching_record.iloc[0]['Site'])
            shifts.append(matching_record.iloc[0]['Shift'])
            managers.append(matching_record.iloc[0]['Leader'])
            employers.append(matching_record.iloc[0]['Employer'])
        else:
            # Handle missing matches
            employee_names.append(row.get('Employee Name'))  # Use 'Name' if available
            employee_ids.append(np.nan)
            work_numbers.append(work_number)
            rags.append(np.nan)
            work_locations.append(np.nan)
            lobs.append(np.nan)
            sites.append(np.nan)
            shifts.append(np.nan)
            managers.append(np.nan)
            employers.append(np.nan)

    # Add new columns to the cleaned dataframe
    cleaned_df['Employee Name'] = employee_names
    cleaned_df['EmployeeID'] = employee_ids
    cleaned_df['WB Work Number'] = work_numbers  # Renamed for consistency
    cleaned_df['RAG'] = rags
    cleaned_df['Work Location'] = work_locations
    cleaned_df['LOB'] = lobs
    cleaned_df['Site'] = sites
    cleaned_df['Shift'] = shifts
    cleaned_df['Manager'] = managers
    cleaned_df['Employer'] = employers

    # Rearrange columns for clarity
    new_column_order = ['Employee Name', 'EmployeeID', 'WB Work Number', 'RAG', 'Work Location',
        'LOB', 'Site', 'Shift', 'Manager', 'Employer'] + date_list
    
    return cleaned_df[new_column_order]

  @staticmethod
  def update_attendance_codes(merged_df, schedule_df):
    """
    Updates attendance codes in the merged dataframe based on the schedule data.

    Args:
        merged_df (pd.DataFrame): The DataFrame containing merged employee and attendance data.
        schedule_df (pd.DataFrame): The DataFrame containing employee schedule information.

    Returns:
        pd.DataFrame: The updated DataFrame with attendance codes potentially including schedule codes.
    """

    # Get a list of columns containing attendance dates
    attendance_date_columns = list(merged_df.columns)[10:]

    # Define a list of possible attendance codes
    attendance_codes = [
        'TRN', 'HD', 'VL', 'ABSA', 'ABSU', 'NCNS', 'RDOT', 'RTWO', 'ATTRIT', 'SL', 
        'EL', 'BL', 'ML', 'OFF', 'SUSPENDED', 'ABSENT-A', 'HalfDay', 'LATE', 
        'ABSENT', 'ABSENT-U', 'ATTRIT', 'CHANGEOFF', 'FLEXI', 'HOLIDAYOFF', 
        'LEAVE', 'RESIGNED', 'SUPPORT', 'TERMINATED', 'TRANSFERINAECALLS', 
        'TRANSFERTOFE', 'TRANSFERED'
    ]

    
    # Iterate through rows in the schedule dataframe
    for index, row in schedule_df.iterrows():
      wb_work_number = row['Work Number']
      # Check if employee ID exists in the merged dataframe
      if wb_work_number in merged_df['WB Work Number'].values:
        for date_column in attendance_date_columns:
          # Check if schedule code exists for the date
          if row[date_column] in attendance_codes:
            # Get existing attendance value for the employee and date
            existing_value = merged_df.loc[(merged_df['WB Work Number'] == wb_work_number), date_column].values[0]
            # Combine existing value with schedule code in parentheses
            updated_value = f"{existing_value} ({row[date_column]})"
            # Update the attendance code in the merged dataframe
            merged_df.loc[(merged_df['WB Work Number'] == wb_work_number), date_column] = updated_value
    return merged_df


  # Function to clean time data in a row (helper function)
  @staticmethod
  def clean_time(row):
      # Apply the cleaning operation to each relevant column in the row
      for col in row.index:
          x = row[col]
          # Ensure x is a string for string operations
          if isinstance(x, str):
              # Attempting to replace "外勤" with an empty string if present and strip whitespace
              x = x.replace("外勤", "").strip()

              # Split the string by newline and extract the first and last times after sorting them
              times = x.split('\n')

              # Checking the conditions and updating x accordingly
              if x.count("\n") > 1 and ":" in x:
                  row[col] = f"{times[0]} - {times[-1]} (MUL)"
              elif x.count("\n") == 1 and x.count(":") == 2:
                  row[col] = f"{times[0]} - {times[-1]}"

              elif x.count(":") == 1:
                  row[col] = f"{x} (MIS)"

              else:
                  # If there's only one time, or no valid time, mark as missing
                  row[col] = x
          else:
              # If x is NaN or not a string, it cannot be cleaned with string methods
              if pd.isna(x):
                  row[col] = np.nan
      return row

  @staticmethod
  def analyze_attendance_time_differences(scheduled_in_time_str, scheduled_out_time_str, actual_in_time_str, actual_out_time_str):
      """
      Analyzes time differences between scheduled and actual attendance times,
      generating attendance status codes.
      """
      try:
          scheduled_in_time = datetime.strptime(scheduled_in_time_str, "%I:%M%p")
          scheduled_out_time = datetime.strptime(scheduled_out_time_str, "%I:%M%p")
          actual_in_time = datetime.strptime(actual_in_time_str, "%H:%M")
          actual_out_time = datetime.strptime(actual_out_time_str, "%H:%M")
      except ValueError as e:
          print(f"Error parsing times - scheduled_in: {scheduled_in_time_str}, scheduled_out: {scheduled_out_time_str}, actual_in: {actual_in_time_str}, actual_out: {actual_out_time_str}")
          raise e
  
      calculate_time_difference = lambda actual, scheduled: (actual - scheduled).total_seconds() / 60
  
      attendance_codes = []
  
      check_in_difference = calculate_time_difference(actual_in_time, scheduled_in_time)
      if check_in_difference <= -16:
          attendance_codes.append("(OT)")
      elif check_in_difference >= 1:
          attendance_codes.append("(L)")
  
      check_out_difference = calculate_time_difference(actual_out_time, scheduled_out_time)
      if check_out_difference >= 16:
          attendance_codes.append("(OT)")
      elif check_out_difference <= -1:
          attendance_codes.append("(L)")
  
      return list(np.unique(attendance_codes))

  @staticmethod
  def merge_final_attendance_codes(self, dataframe_with_codes, schedule_dataframe):
    """
    Incorporates final attendance codes based on schedule information.

    Args:
        dataframe_with_codes (pd.DataFrame): DataFrame with existing attendance codes.
        schedule_dataframe (pd.DataFrame): DataFrame containing schedule data.

    Returns:
        pd.DataFrame: DataFrame with final attendance codes applied.
    """

    attendance_date_columns = list(dataframe_with_codes.columns)[10:]

    for index, schedule_row in schedule_dataframe.iterrows():
      employee_work_number = schedule_row['Work Number']

      if employee_work_number in dataframe_with_codes['WB Work Number'].values:
        for date_column in attendance_date_columns:
          schedule_value = schedule_row[date_column]
          attendance_value = dataframe_with_codes.loc[dataframe_with_codes['WB Work Number'] == employee_work_number, date_column].values[0]

          # Check for absence based on schedule
          if pd.isnull(attendance_value) and not pd.isnull(schedule_value) and isinstance(schedule_value, str):
            # Mark as absent if scheduled but no attendance
            new_value = f"{attendance_value} (ABSENT)" if pd.isnull(attendance_value) else attendance_value
            dataframe_with_codes.loc[dataframe_with_codes['WB Work Number'] == employee_work_number, date_column] = new_value

          # Apply attendance time-based codes
          elif isinstance(schedule_value, str) and isinstance(attendance_value, str) and schedule_value.count(":") == 2 and attendance_value.count(":") == 2:
            # Extract time information from strings
            scheduled_in_time_str = schedule_value[0:7]
            scheduled_out_time_str = schedule_value[8:15]
            actual_in_time_str = attendance_value[0:5]
            actual_out_time_str = attendance_value[6:11]

            # Analyze time differences and retrieve codes
            applicable_codes = self.analyze_attendance_time_differences(scheduled_in_time_str, scheduled_out_time_str, actual_in_time_str, actual_out_time_str)

            # Update value with applicable codes
            if applicable_codes:
              codes_string = ' '.join(applicable_codes)
              new_value = f"{attendance_value} {codes_string}"
            else:
              new_value = attendance_value  # Keep existing value if no codes

            dataframe_with_codes.loc[dataframe_with_codes['WB Work Number'] == employee_work_number, date_column] = new_value

    return dataframe_with_codes

  @staticmethod
  def transform_attendance_data(attendance_dataframe):
    """
    Reformats attendance data into a DataFrame with separate columns for employee details,
    site, shift, manager, dates, time in, time out, and remarks.

    Args:
        attendance_dataframe (pd.DataFrame): DataFrame containing raw attendance data.

    Returns:
        pd.DataFrame: DataFrame with transformed and structured attendance information.
    """

    # Prepare a list to hold the formatted data
    formatted_records = []

    # Regular expression to match time and remarks in parentheses
    time_pattern = re.compile(r'(\d{2}:\d{2})(-(\d{2}:\d{2}))?(\(.*\))?')

    # Extract a list of dates from column names (excluding the last two columns)
    date_columns = attendance_dataframe.columns[10:]
    attendance_dates = [pd.to_datetime(date).date() for date in date_columns]

    # Iterate through each row in the DataFrame
    for index, data_row in attendance_dataframe.iterrows():
      # Extract employee details
      employee_name = data_row['Employee Name']
      employee_id = data_row['EmployeeID']
      work_number = data_row['WB Work Number']
      lob = data_row['LOB']
      employee_site = data_row['Site']
      employee_shift = data_row['Shift']
      team_leader = data_row['Manager']

      # Iterate over each date to construct individual records
      for attendance_date in attendance_dates:
        # Initialize placeholders for time and remarks
        in_time, out_time, comments = None, None, None

        # Extract time data from the current date's cell
        time_data = data_row[str(attendance_date)]

        # Process time data if it's not missing
        if pd.notna(time_data):
          time_string = str(time_data).strip()
          match = time_pattern.match(time_string)
          if match:
            in_time = match.group(1)
            out_time = match.group(3) if match.group(3) else None
            comments = match.group(4) if match.group(4) else None
          else:
            # If format doesn't match, consider entire string as a comment
            comments = time_string

        # Append the structured data to the formatted records list
        formatted_records.append({
            'Date': attendance_date,
            'Employee Name': employee_name,
            'Employee ID': employee_id,
            'WB Work Number': work_number,
            'LOB': lob,
            'Site': employee_site,
            'Shift': employee_shift,
            'Manager': team_leader,
            'Time In': in_time,
            'Time Out': out_time,
            'Remarks': comments
        })

    # Convert the formatted records list into a DataFrame
    formatted_dataframe = pd.DataFrame(formatted_records)

    return formatted_dataframe

class AnalysisUtils:
  @staticmethod
  def metric_count(data_frame):
    """Counts the occurrences of specific codes in the 'Remarks' column of a DataFrame.

    Args:
        data_frame (pandas.DataFrame): The DataFrame containing manager and code information.

    Returns:
        tuple: A tuple containing the counts of 'MIS', 'MUL', 'ABSENT', and 'L' code occurrences.
    """

    missing_information_count = data_frame['Remarks'].str.contains('(MIS)', case=False).sum()
    multiple_occurrence_count = data_frame['Remarks'].str.contains('(MUL)', case=False).sum()
    absent_count = data_frame['Remarks'].str.contains('(ABSENT)', case=False).sum()
    late_count = data_frame['Remarks'].str.contains('(L)', case=False).sum()

    return missing_information_count, multiple_occurrence_count, absent_count, late_count


  @staticmethod
  def count_codes_per_manager(data_frame):
      """Counts the occurrences of specific codes for each manager in the DataFrame.

      Args:
          data_frame (pandas.DataFrame): The DataFrame containing manager and code information.

      Returns:
          pandas.DataFrame: A DataFrame with columns for manager, code status, and count.
      """

      managers = []
      code_statuses = []
      code_counts = []

      for manager_name in data_frame['Manager'].unique():
          for code_status in ['(ABSENT)', '(MUL)', '(MIS)', '(L)']:
              count_of_occurrences = data_frame.loc[
                  data_frame['Manager'] == manager_name
              ]['Remarks'].str.contains(code_status).sum()
              managers.append(manager_name)
              code_statuses.append(code_status)
              code_counts.append(count_of_occurrences)

      summary_data_frame = pd.DataFrame({
          'Manager': managers,
          'Status': code_statuses,
          'Count': code_counts
      })

      return summary_data_frame
  
  @staticmethod
  def count_code_per_date(df):

    date_list = []
    code_list = []
    code_counts = []
    for date in df['Date'].unique():
      for code_remarks in ['(ABSENT)', '(MUL)', '(MIS)', '(L)']:
        count = df.loc[df['Date'] == date]['Remarks'].str.contains(code_remarks).sum()
        date_list.append(date)
        code_list.append(code_remarks)
        code_counts.append(count)

    count_df = pd.DataFrame({
        'Date': date_list,
        'Status': code_list,
        'Count': code_counts
    })

    return count_df


  @staticmethod
  def plot_code_occurrence_by_date(data_frame, code_status):
    """Plots the daily occurrences of a specific code status in a DataFrame.

    Args:
        data_frame (pandas.DataFrame): The DataFrame containing code status and date information.
        code_status (str): The specific code status to plot (e.g., '(L)' for Late).

    Returns:
        plotly.graph_objects.Figure: A plotly figure representing the code occurrences over time.
    """

    code_descriptions = {
        '(L)': 'Late',
        '(MUL)': 'Multiple Punches',
        '(MIS)': 'Missed Punch',
        '(ABSENT)': 'Absent'
    }

    peak_count = data_frame.loc[data_frame['Status'] == code_status]['Count'].max()
    peak_date = data_frame.loc[data_frame['Count'] == peak_count]['Date'].values[0]

    figure = px.line(
        data_frame.loc[data_frame['Status'] == code_status],
        x="Date",
        y="Count",
        markers=True,
        title=f"When do the most {code_descriptions[code_status]} entries occur?<br>"
              f"<sup>Peak: {peak_count} occurrences in {peak_date}</sup>"
    )

    return figure
  
  @staticmethod
  def plot_leaders_by_code_occurrence(data_frame, code_status):
      """Plots a bar chart showing leaders with the most occurrences of a specific code.

      Args:
          data_frame (pandas.DataFrame): The DataFrame containing code status, manager, and count information.
          code_status (str): The specific code status to plot (e.g., '(L)' for Late).

      Returns:
          plotly.graph_objects.Figure: A plotly figure representing the leader distribution for the code.
      """

      code_descriptions = {
          '(L)': 'Late',
          '(MUL)': 'Multiple Punches',
          '(MIS)': 'Missed Punch',
          '(ABSENT)': 'Absent'
      }

      filtered_data = data_frame.loc[data_frame['Status'] == code_status]
      sorted_data = filtered_data.sort_values(by='Count', ascending=True)

      figure = px.bar(
          sorted_data,
          x="Count",
          y="Manager",
          orientation='h',
          title=f"Who are responsible for records of {code_descriptions[code_status]}?<br>"
                f"<sup>Leaders with the Most {code_descriptions[code_status]} Employees</sup>"
      )

      return figure

