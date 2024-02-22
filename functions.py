import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore", message="Workbook contains no default style")
from PIL import Image
import streamlit as st

pd.options.mode.chained_assignment = None  # Set pandas option to suppress chained assignment warning

im = Image.open("images/neusoft_logo.png")
master_list = pd.read_csv('master_list.csv')

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
def process_attendance_data(df) -> pd.DataFrame:
    """
    Process attendance data from an Excel file.

    Args:
        df (pd.DataFrame): Raw DataFrame

    Returns:
        pd.DataFrame: A DataFrame containing the processed data.
    """

    # Generate a list of dates from the dataframe
    generated_dates = create_list_of_dates(df)

    # Define column headers including the generated dates
    columns = ["Name", "Attendance Group", "Department", "WB Work Number", "Position", "User ID"]
    columns.extend(generated_dates)

    # Set the column headers of the dataframe
    df.columns = columns

    # Get the cleaned dataframe by removing the first two rows
    cleaned_df = df[2:]

    # Process each row in the dataframe starting from the 7th column (index 6)
    for index, row in cleaned_df.iloc[:, 6:].iterrows():
        # Process each cell in the row
        for col in row.index:
            # Check if the cell value is a string
            if isinstance(row[col], str):
                # Split the string by newline character and remove leading/trailing whitespaces
                split_row = [val.strip() for val in row[col].split("\n")]
                # Check if there are at least two elements and the first and last elements are different
                if len(split_row) >= 2 and split_row[0] != split_row[-1]:
                    # If conditions met, set the cell value to a range formed by the first and last elements
                    new_value = f"{split_row[0]} - {split_row[-1]}"
                else:
                    # If conditions not met, set the cell value to "Missed Punch"
                    new_value = f"Missed Punch ({row[col].strip()})"
            # Check if the cell value is a float and the corresponding date is Saturday or Sunday
            elif isinstance(row[col], float) and datetime.strptime(col, '%Y-%m-%d').strftime("%A") in ['Saturday', 'Sunday']:
                # If the date is weekend, set the cell value to "OFF"
                new_value = 'OFF'
            else:
                # If none of the above conditions are met, set the cell value to "No Log"
                new_value = "No Log"
            # Update the cell value in the dataframe
            cleaned_df.at[index, col] = new_value

    # Define the columns to select
    columns_to_select = ['Name', 'WB Work Number']
    # Add the remaining columns (from the 7th column onwards) to the list
    columns_to_select.extend(cleaned_df.iloc[:, 6:].columns)

    # Select the desired columns from the cleaned dataframe
    result_df = cleaned_df.loc[:, columns_to_select]

    return result_df

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
    rag_column = []
    lob_column = []
    site_column = []
    shift_column = []

    # Iterate through each row of the cleaned dataframe
    for index, row in cleaned_df.iterrows():
        # Check if the 'WB Work Number' exists in the master list
        if row['WB Work Number'] in master_list['WB Work Number'].values:
            # Filter the master list based on matching 'WB Work Number'
            value = master_list[master_list['WB Work Number'] == row['WB Work Number']]
            # Extract and append data from the master list to respective columns
            employee_name_column.append(value['Employee Name'].values[0])
            rag_column.append(value['RAG'].values[0])
            lob_column.append(value['LOB'].values[0])
            shift_column.append(value['Shift'].values[0])
            site_column.append(value['Site'].values[0])
        else:
            # If 'WB Work Number' not found, use data from the current row and fill missing values with NaN
            employee_name_column.append(row['Name'])
            rag_column.append(np.nan)
            lob_column.append(np.nan)
            shift_column.append(np.nan)
            site_column.append(np.nan)

    # Add new columns to the cleaned dataframe
    cleaned_df['Employee Name'] = employee_name_column
    cleaned_df['RAG'] = rag_column
    cleaned_df['LOB'] = lob_column
    cleaned_df['Shift'] = shift_column
    cleaned_df['Site'] = site_column

    # Define the columns to keep
    columns_to_select = ['Employee Name', 'LOB', 'Shift', 'Site']

    # Add the remaining columns (from the 7th column onwards) to the selection list
    columns_to_select.extend(cleaned_df.iloc[:, 2:].columns)

    # Select the desired columns from the dataframe
    cleaned_df = cleaned_df[columns_to_select]

    # Remove the last 5 columns from the selected dataframe
    cleaned_df = cleaned_df.iloc[:, :-5]

    return cleaned_df


def apply_combined_styles(val):
        """
        Apply background color to cells based on multiple conditions.

        Args:
            val (str or int): The value to check for coloring.

        Returns:
            str: A CSS style string defining the background color for the cell.

        """

        # Check for 'No Log' or 'OFF' values
        if isinstance(val, str):

            if len(val.split('-')) == 2 and val[0].strip().isdigit():
                start_time_str = val.split('-')[0].strip()
                end_time_str = val.split('-')[1].strip()

                # Convert start and end times to datetime objects
                start_time = datetime.strptime(start_time_str, '%H:%M')
                end_time = datetime.strptime(end_time_str, '%H:%M')

                # Convert start and end times to datetime objects
                start_time = datetime.strptime(start_time_str, '%H:%M')
                end_time = datetime.strptime(end_time_str, '%H:%M')

                # Calculate duration
                duration = end_time - start_time

                # Extract hours and minutes
                hours_worked = duration.seconds // 3600
                minutes_worked = (duration.seconds % 3600) // 60

                if hours_worked <= 6:
                    background_color = '#fff3aa'
                    return f'background-color: {background_color}'

            elif 'Missed Punch' in val:
                background_color = '#ffd0a8'
                return f'background-color: {background_color}'

            elif val == 'No Log':
                background_color = '#ffb1b1'
                return f'background-color: {background_color}'

            elif val == 'OFF':
                background_color = '#e4fbdc'
                return f'background-color: {background_color}'
            else:
                pass

def count_no_log(x):
    return (x == 'No Log').sum()

def count_missed_punch(x):
    # Check if the value is a string and if 'Missed Punch' is present
    if isinstance(x, str):
        return ('Missed Punch' in x)
    else:
        return False  # Return False for non-string values

def count_logs(df):
    # Apply count_no_log function to relevant columns and assign result to 'No Log Count' column
    df['No Log Count'] = df.iloc[:, 5:].apply(count_no_log, axis=1)
    
    # Apply count_missed_punch function to relevant columns and assign result to 'Missed Punch Count' column
    df['Missed Punch Count'] = df.iloc[:, 5:].applymap(count_missed_punch).sum(axis=1)
    
    # Sort and return the DataFrames
    no_log_sorted = df[['Employee Name', 'No Log Count']].sort_values(by='No Log Count', ascending=False)
    missed_punch_sorted = df[['Employee Name', 'Missed Punch Count']].sort_values(by='Missed Punch Count', ascending=False)
    
    return no_log_sorted, missed_punch_sorted


def find_duplicates_by_wb_work_number(df):
    # Drop rows with null values in 'WB Work Number' column
    df_cleaned = df.dropna(subset=['WB Work Number'])
    
    # Identify duplicated rows based on 'WB Work Number'
    duplicated_rows = df_cleaned[df_cleaned.duplicated(subset='WB Work Number', keep=False)]
    
    return duplicated_rows[['Name', 'WB Work Number']]