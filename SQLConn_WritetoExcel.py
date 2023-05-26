import pyodbc
import pandas as pd
from dateutil import rrule
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Set up the connection string
server_name = 'chin2'
db_name = 'HomeDB'
connection_string = f'DRIVER={{SQL Server}};SERVER={server_name};DATABASE={db_name};Trusted_Connection=yes;'

# Establish the database connection
connection = pyodbc.connect(connection_string)

try:
    # Create a cursor object to interact with the database
    cursor = connection.cursor()

    # Define the stored procedure execution query
    stored_procedure = 'EXEC sample_procedure @input_date=?'

    # Define the date range
    start_date = pd.to_datetime('2022-01-01')
    end_date = pd.to_datetime('2023-03-31')

    # Create an Excel workbook
    workbook = Workbook()

    # Iterate over each date in the range
    for dt in rrule.rrule(rrule.DAILY, dtstart=start_date, until=end_date):
        # Format the date as YYYY-MM-DD
        date_str = dt.strftime('%Y-%m-%d')

        # Set the parameter value
        date_parameter = date_str

        # Execute the stored procedure
        cursor.execute(stored_procedure, date_parameter)

        # Fetch all the rows returned by the stored procedure
        results = cursor.fetchall()

        if results:
            # Convert the results to a list of strings
            results_strings = [tuple(str(value) for value in row) for row in results]

            # Create a DataFrame with the fetched results
            df = pd.DataFrame(results_strings, columns=[column[0] for column in cursor.description])

            # Create a sheet in the workbook with the date as the name
            sheet = workbook.create_sheet(title=date_str)

            # Write the DataFrame to the sheet
            for r in dataframe_to_rows(df, index=False, header=True):
                sheet.append(r)

    # Remove the default sheet created by openpyxl
    workbook.remove(workbook['Sheet'])

    # Specify the Excel file path
    excel_file_path = 'output.xlsx'

    # Save the workbook to Excel
    workbook.save(excel_file_path)

    print(f'Data saved to {excel_file_path} successfully.')

except pyodbc.Error as e:
    # Handle any errors that may occur during execution
    print('Error executing stored procedure:', str(e))

finally:
    # Close the cursor and connection
    cursor.close()
    connection.close()
