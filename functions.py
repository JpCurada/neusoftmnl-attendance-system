import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import warnings
warnings.filterwarnings("ignore", message="Workbook contains no default style")
from PIL import Image
import streamlit as st


pd.options.mode.chained_assignment = None  # Set pandas option to suppress chained assignment warning

im = Image.open("images/neusoft_logo.png")

def process_masterlist(raw_masterlist):
  df_act = pd.read_excel(raw_masterlist, sheet_name='Active')
  df_in = pd.read_excel(raw_masterlist, sheet_name='Inactive')
  concat_df = pd.concat([df_act, df_in], ignore_index=True)
  df = concat_df[['Employee Name', 'WB Work Number','RAG', 'Work Location', 'Shift', 'Site', 'LOB', 'Employer']]
  return df

def create_list_of_dates(df):
    """
    Create a list of dates from the given DataFrame.

    Parameters:
        df (DataFrame): The DataFrame containing dates.

    Returns:
        list: A list of dates generated from the DataFrame.
    """
    # Extract text from the first column of the DataFrame
    text = df.columns[0]
    # Split the text to extract start and end dates
    dates = text.split("：")[1].split(" 至 ")
    # Extract start and end dates
    start_date = dates[0]
    end_date = dates[1]

    # Convert start and end dates to datetime objects
    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")

    # Generate dates between start and end dates
    dates_generated = []
    while start_date <= end_date:
        # Append formatted date to the list
        dates_generated.append(start_date.strftime("%Y-%m-%d"))
        # Increment start date by one day
        start_date += timedelta(days=1)

    return dates_generated

@st.cache_data
def prep_attendance_data(df):

    generated_dates = create_list_of_dates(df)

    sliced_df = df[2:]
    sliced_df = sliced_df.copy()

    columns = ["Name", "Attendance Group", "Department", "WB Work Number", "Position", "User ID"]
    columns.extend(generated_dates)

    sliced_df.columns = columns

    return sliced_df

@st.cache_data
def process_masterlist(raw_masterlist):
  df_act = pd.read_excel(raw_masterlist, sheet_name='Active')
  df_in = pd.read_excel(raw_masterlist, sheet_name='Inactive')
  concat_df = pd.concat([df_act, df_in], ignore_index=True)
  df = concat_df[['Employee Name', 'Employee Code (ID)' ,'WB Work Number','RAG', 'Work Location', 'Shift', 'Site', 'LOB', 'Leader','Employer']]
  return df

@st.cache_data
def merge_master_list(cleaned_df: pd.DataFrame, master_list: pd.DataFrame) -> pd.DataFrame:
    """
    Merge data from the master list into the cleaned dataframe.

    Parameters:
        cleaned_df (pd.DataFrame): The DataFrame to be cleaned.
        master_list (pd.DataFrame): The master list DataFrame containing additional data.

    Returns:
        pd.DataFrame: The cleaned DataFrame with merged data.
    """

    # Initialize empty lists to store data for new columns
    employee_name_column = []
    employee_id_column = []
    wb_work_number = []  # Corrected variable name from wb_work_numbe to wb_work_number
    rag_column = []
    work_location = []
    lob_column = []
    site_column = []
    shift_column = []
    leader_column = []
    employer_column = []

    # Iterate through each row of the cleaned dataframe
    for index, row in cleaned_df.iterrows():
        # Check if the 'WB Work Number' exists in the master list
        if row['WB Work Number'] in master_list['WB Work Number'].values:
            # Filter the master list based on matching 'WB Work Number'
            value = master_list[master_list['WB Work Number'] == row['WB Work Number']].iloc[0]

            # Extract and append data from the master list to respective columns
            employee_name_column.append(value['Employee Name'])
            employee_id_column.append(value['Employee Code (ID)'])
            wb_work_number.append(row['WB Work Number'])  # Assuming you want to keep the original WB Work Number
            rag_column.append(value['RAG'])
            work_location.append(value['Work Location'])
            lob_column.append(value['LOB'])
            site_column.append(value['Site'])
            shift_column.append(value['Shift'])
            leader_column.append(value['Leader'])
            employer_column.append(value['Employer'])
        else:
            # If 'WB Work Number' not found, append NaNs or placeholders as appropriate
            employee_name_column.append(row['Name'])  # Assuming 'Name' exists in cleaned_df
            employee_id_column.append(np.nan)
            wb_work_number.append(row['WB Work Number'])  # Keeping the original WB Work Number
            rag_column.append(np.nan)
            work_location.append(np.nan)
            lob_column.append(np.nan)
            site_column.append(np.nan)
            shift_column.append(np.nan)
            leader_column.append(np.nan)
            employer_column.append(np.nan)

    # Add new columns to the cleaned dataframe
    cleaned_df['Employee Name'] = employee_name_column
    cleaned_df['EmployeeID'] = employee_id_column
    cleaned_df['WB Work Number'] = wb_work_number  # Ensure this is correct according to your dataframe structure
    cleaned_df['RAG'] = rag_column
    cleaned_df['Work Location'] = work_location
    cleaned_df['LOB'] = lob_column
    cleaned_df['Shift'] = shift_column
    cleaned_df['Site'] = site_column
    cleaned_df['Leader'] = leader_column
    cleaned_df['Employer'] = employer_column

    # Select the last 9 columns
    last_nine_columns = cleaned_df.iloc[:, -9:]

    # Select the columns you want to keep (excluding the first 6 and the last 9 which are being moved)
    remaining_columns = cleaned_df.iloc[:, 6:-9]

    # Concatenate the last 9 columns with the remaining columns
    new_df = pd.concat([last_nine_columns, remaining_columns], axis=1)

    return new_df

def process_schedule_data(file_path):
    schedule_dates = pd.read_excel(file_path, sheet_name='RBC').iloc[0].values[6:]
    formatted_dates = [date.strftime('%Y-%m-%d') for date in schedule_dates]
    columns = ['i','No.','LOB','EmployeeID', 'WBWorkNumber', 'Name']
    columns.extend(formatted_dates)

    rbc_df = pd.read_excel(file_path, sheet_name='RBC')
    sliced_rbc = rbc_df[3:]
    hsq_df = pd.read_excel(file_path, sheet_name='HSQ')
    sliced_hsq = hsq_df[3:]
    idn_df = pd.read_excel(file_path, sheet_name='IDN')
    sliced_idn = idn_df[3:]
    isa_df = pd.read_excel(file_path, sheet_name='ISA')
    sliced_isa = isa_df[3:]

    df = pd.concat([sliced_rbc, sliced_hsq, sliced_idn, sliced_isa], ignore_index=True)
    df.columns = columns

    for col in columns[6:]:
      df[col] = df[col].str.replace(r'\s+', '', regex=True)

    return df

def apply_codes(merged_df, sched_df):

  attendance_date_cols = list(merged_df.columns)[9:]
  codes = ['TRN', 'HD','VL', 'ABSA', 'ABSU', 'NCNS', 'RDOT', 'RTWO', 'ATTRIT', 'SL', 'EL', 'BL', 'ML']

  for index, row in sched_df.iterrows():
    for date_col in attendance_date_cols:
      if row['EmployeeID'] in merged_df['EmployeeID'].values:
        if row[date_col] in codes:
          existing_val = merged_df.loc[(merged_df['EmployeeID'] == row['EmployeeID']), date_col].values[0]
          updated_val = f"{existing_val} ({row[date_col]})"
          merged_df.loc[(merged_df['EmployeeID'] == row['EmployeeID']), date_col] = updated_val

  applied_codes_df = merged_df

  return applied_codes_df

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

def compute_duration(start_time, end_time):
    return end_time - start_time

def operate_time(sched_in_str, sched_out_str, actual_in_str, actual_out_str):

  # Parsing input strings into datetime objects
  sched_in = datetime.strptime(sched_in_str, "%I:%M%p")
  sched_out = datetime.strptime(sched_out_str, "%I:%M%p")
  actual_in = datetime.strptime(actual_in_str, "%H:%M")
  actual_out = datetime.strptime(actual_out_str, "%H:%M")
  
  # Calculating differences in minutes
  diff_in_minutes = lambda actual, sched: (actual - sched).total_seconds() / 60
  
  codes = []
  
  # Analyzing check-in status
  if diff_in_minutes(actual_in, sched_in) <= -16:
      codes.append("(OT)")
  elif diff_in_minutes(actual_in, sched_in) >= 1:
      codes.append("(L)")
  
  # Analyzing check-out status
  if diff_in_minutes(actual_out, sched_out) >= 16:
      codes.append("(OT)")
  elif diff_in_minutes(actual_out, sched_out) <= -1:
      codes.append("(L)")

  return codes


def add_final_codes(applied_codes_df, sched_df):

  attendance_date_cols = list(applied_codes_df.columns)[9:]

  for index, sched_row in sched_df.iterrows():
      for date_col in attendance_date_cols:
          if sched_row['EmployeeID'] in applied_codes_df['EmployeeID'].values:
              sched_val = sched_row[date_col]
              attdn_val = applied_codes_df.loc[applied_codes_df['EmployeeID'] == sched_row['EmployeeID'], date_col].values[0]

              # Check if attdn_val is not null and sched_val indicates presence or is scheduled
              if pd.isnull(attdn_val) and not pd.isnull(sched_val) and isinstance(sched_val, str):
                  # Mark as absent if there is a schedule but no attendance
                  new_val = f"{attdn_val} (ABSENT)" if pd.isnull(attdn_val) else attdn_val
                  applied_codes_df.loc[applied_codes_df['EmployeeID'] == sched_row['EmployeeID'], date_col] = new_val

              # If both sched_val and attdn_val are properly formatted time strings
              elif isinstance(sched_val, str) and isinstance(attdn_val, str) and sched_val.count(":") == 2 and attdn_val.count(":") == 2:
                  # Convert to datetime objects for time comparisons

                  sched_in_time_dt = datetime.strptime(sched_val[0:7], '%I:%M%p').strftime('%H:%M')
                  sched_out_time_dt = datetime.strptime(sched_val[8:15], '%I:%M%p').strftime('%H:%M')
                  actual_in_time_dt = datetime.strptime(attdn_val[0:5], "%H:%M").strftime('%H:%M')
                  actual_out_time_dt = datetime.strptime(attdn_val[6:11], "%H:%M").strftime('%H:%M')

                  # Call operate_time with datetime objects
                  applicable_codes = operate_time(sched_in_time_dt, sched_out_time_dt, actual_in_time_dt, actual_out_time_dt)

                  # Update only with applicable codes
                  if applicable_codes:
                      codes_str = ' '.join(applicable_codes)  # Convert list of codes to a single string
                      new_val = f"{attdn_val} {codes_str}"
                  else:
                      new_val = attdn_val  # If no codes, keep existing value

                  applied_codes_df.loc[applied_codes_df['EmployeeID'] == sched_row['EmployeeID'], date_col] = new_val
  return applied_codes_df

def apply_color(x):
    # Initialize color as None for cases where x doesn't match any condition
    color = None

    # Ensure x is a string to safely use 'in' operation
    if isinstance(x, str):
        # Check for combinations of conditions first
        if 'MUL' in x and 'OT' in x and 'L' in x:
            color = "#A93226"
        elif 'MUL' in x and 'UT' in x and 'L' in x:
            color = "#E74C3C"
        elif 'UT' in x and 'L' in x:
            color = "#EC7063"
        elif 'OT' in x and 'L' in x:
            color = "Violet"
        # Check for single conditions
        elif 'ABSENT' in x:
            color = "Red"
        elif 'L' in x:
            color = "Yellow"
        elif 'MUL' in x:
            color = "Gray"
        elif 'MIS' in x:
            color = "Pink"
        elif 'OT' in x:
            color = "Blue"
        elif 'UT' in x:
            color = "Orange"
        # Return the appropriate style string based on the color
        if color:
            return f'background-color: {color}'
        else:
            return ''

def reformat_df(attendance_df):

  # Prepare a list to hold the formatted data
  formatted_data = []

  # Regular expression to match time and remarks in parentheses
  time_re = re.compile(r'(\d{2}:\d{2})(-(\d{2}:\d{2}))?(\(.*\))?')

  # Iterate over each row in the DataFrame
  for index, row in attendance_df.iterrows():
      # Extract the employee's name and ID which are constant across the dates
      employee_name = row['Employee Name']
      employee_id = row['EmployeeID']
      
      # Iterate over each date column to construct individual records
      for date_column in attendance_df.columns[9:-2]:  # Adjusted to exclude the last two columns
          # Extract the date from the column name
          date = pd.to_datetime(date_column).date()
          # Extract the time information from the cell
          time_data = row[date_column]
          
          # Check if the time data is not NaN
          if pd.notna(time_data):
              time_string = str(time_data).strip()
              match = time_re.match(time_string)
              if match:
                  time_in = match.group(1)
                  time_out = match.group(3) if match.group(3) else None
                  remarks = match.group(4) if match.group(4) else None
              else:
                  # If the format does not match, log the entire string as a remark
                  time_in, time_out, remarks = None, None, time_string

              # Append the data to the formatted data list
              formatted_data.append({
                  'Date': date,
                  'Employee Name': employee_name,
                  'Employee ID': employee_id,
                  'Time In': time_in,
                  'Time Out': time_out,
                  'Remarks': remarks
              })

  # Convert the formatted data into a DataFrame
  formatted_df = pd.DataFrame(formatted_data)

  return formatted_df





def count_multiple_log(x):
    if isinstance(x, str):
        return ('(MUL)' in x)
    else:
        return False  # Return False for non-string values

def count_missed_punch(x):
    if isinstance(x, str):
        return ('(MIS)' in x)
    else:
        return False  # Return False for non-string values

def count_logs(df):
    # Apply count_no_log function to relevant columns and assign result to 'No Log Count' column
    df['Multiple Logs Count'] = df.iloc[:, 9:].applymap(count_multiple_log).sum(axis=1)
    
    # Apply count_missed_punch function to relevant columns and assign result to 'Missed Punch Count' column
    df['Missed Punch Count'] = df.iloc[:, 9:].applymap(count_missed_punch).sum(axis=1)
    
    # Sort and return the DataFrames
    mul_sorted = df[['Employee Name', 'Multiple Logs Count']].sort_values(by='Multiple Logs Count', ascending=False)
    mis_sorted = df[['Employee Name', 'Missed Punch Count']].sort_values(by='Missed Punch Count', ascending=False)
    
    return mul_sorted, mis_sorted


def find_duplicates_by_wb_work_number(df):
    # Drop rows with null values in 'WB Work Number' column
    df_cleaned = df.dropna(subset=['WBWorkNumber'])
    
    # Identify duplicated rows based on 'WB Work Number'
    duplicated_rows = df_cleaned[df_cleaned.duplicated(subset='WBWorkNumber', keep=False)]
    
    return duplicated_rows[['Name', 'WBWorkNumber']]
