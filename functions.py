import pandas as pd  # Importing pandas library for data manipulation
from datetime import datetime, timedelta
from PIL import Image

pd.options.mode.chained_assignment = None  # Set pandas option to suppress chained assignment warning

im = Image.open("images/neusoft_logo.png")

def create_list_of_dates(df):
    """
    Extracts the start and end dates from the DataFrame column and generates a list of dates.

    Args:
        df (DataFrame): The DataFrame containing the date range.

    Returns:
        list: A list of dates between the start and end dates.
    """
    text = df.columns[0]  # Extracting the date range text from the DataFrame column
    dates = text.split("：")[1].split(" 至 ")  # Splitting the text to extract start and end dates
    start_date = dates[0]  # Start date extracted from the split text
    end_date = dates[1]  # End date extracted from the split text

    start_date = datetime.strptime(start_date, "%Y-%m-%d")  # Converting start date string to datetime object
    end_date = datetime.strptime(end_date, "%Y-%m-%d")  # Converting end date string to datetime object

    dates_generated = []  # Initializing list to store generated dates

    # Generating dates between start and end dates
    while start_date <= end_date:
        dates_generated.append(start_date.strftime("%Y-%m-%d"))  # Appending formatted date to the list
        start_date += timedelta(days=1)  # Incrementing start date by one day

    return dates_generated  # Returning the list of generated dates

def clean_process_create_df(df, generated_dates):
    """
    Cleans the DataFrame, processes it, and creates a new DataFrame.

    Args:
        df (DataFrame): The DataFrame containing the data to be cleaned and processed.
        generated_dates (list): A list of dates.

    Returns:
        DataFrame: The new DataFrame after cleaning, processing, and creating.
    """
    # Clean DataFrame and apply appropriate column headers
    num_of_days = [datetime.strptime(date, '%Y-%m-%d').strftime("%d") for date in generated_dates]  # Extracting day numbers from dates
    columns = ["Name", "Attendance Group", "Department", "Employee Number", "Position", "User ID"]  # Defining column headers
    columns.extend(num_of_days)  # Extending column headers with day numbers
    df.columns = columns  # Assigning new column headers to the DataFrame
    cleaned_df = df[2:]  # Removing unnecessary rows from the DataFrame

    # Process cleaned DataFrame
    for index, row in cleaned_df.iloc[:, 6:].iterrows():  # Iterating over rows starting from the 7th column
        for col in row.index:  # Iterating over column indices
            if isinstance(row[col], str):  # Checking if the cell value is a string
                row[col] = row[col[.replace('外勤', '') # remove the 'field' (system typo(
                split_row = [val.strip() for val in row[col].split("\n")]  # Splitting the string value by newline and stripping whitespaces
                if len(split_row) >= 2 and split_row[0] != split_row[-1]:  # Checking if there are multiple values and they are different
                    new_value = f"{split_row[0]} - {split_row[-1]}"  # Combining multiple values with a hyphen
                else:
                    new_value = f"Missed Punch ({row[col].strip()})"  # Assigning "Missed Punch" if there's only one value or it's the same
            else:
                new_value = "No Log"  # Assigning "No Log" if the cell value is not a string
            cleaned_df.at[index, col] = new_value  # Assigning the new value to the DataFrame cell

    # Create new DataFrame from processed DataFrame
    new_columns = list(cleaned_df['Name'].unique())  # Extracting unique names from the "Name" column
    new_columns.insert(0, "Date")  # Inserting "Date" column at the beginning
    new_columns.insert(1, "Weekday")  # Inserting "Weekday" column after "Date"
    new_df = pd.DataFrame(columns=new_columns)  # Creating a new DataFrame with specified columns

    new_df["Date"] = generated_dates  # Assigning generated dates to the "Date" column
    new_df["Weekday"] = [datetime.strptime(date, '%Y-%m-%d').strftime("%A") for date in generated_dates]  # Extracting weekdays from dates

    for name in new_df.columns[1:]:  # Iterating over columns starting from the 2nd column
        if name in cleaned_df["Name"].values:  # Checking if the name exists in the cleaned DataFrame
            name_data = cleaned_df[cleaned_df["Name"] == name].iloc[:, 6:].values[0]  # Extracting attendance data for the name
            new_df[name] = name_data  # Assigning attendance data to the corresponding column in the new DataFrame

    for index, row in new_df.iterrows():  # Iterating over rows in the new DataFrame
        for col in row.index:  # Iterating over column indices
            if row[col] == "No Log" and row["Weekday"] in ["Saturday", "Sunday"]:  # Checking for "No Log" entries on weekends
                new_value = "OFF"  # Assigning "OFF" for weekends with no log
                new_df.at[index, col] = new_value  # Assigning the new value to the DataFrame cell

    return new_df  # Returning the new DataFrame

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

            if len(val.split('-')) == 2:
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
