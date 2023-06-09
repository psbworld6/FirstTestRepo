import os
from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime

# Directory path containing the Excel files
directory_path = 'path/to/directory'  # Replace with the actual directory path

# Destination file path
destination_file_path = 'path/to/destination/file.xlsx'  # Replace with the actual destination file path

# Create a new workbook for the consolidated reports
consolidated_workbook = Workbook()

# Create a list to store the file paths and their corresponding month and year
file_info_list = []

# Iterate over the files in the directory
for filename in os.listdir(directory_path):
    if filename.endswith(".xlsx"):
        file_path = os.path.join(directory_path, filename)
        
        # Get the month and year information from the file name
        month_year = datetime.strptime(filename.split('.')[0], '%B%Y')

        # Add the file path and its corresponding month and year to the list
        file_info_list.append((file_path, month_year))

# Sort the file paths based on the month and year
sorted_file_info_list = sorted(file_info_list, key=lambda x: x[1])

# Iterate over the sorted file paths
for file_path, month_year in sorted_file_info_list:
    # Load the source workbook
    source_workbook = load_workbook(file_path, read_only=True)

    # Check if the "Consolidated Test Report" sheet exists
    if "Consolidated Test Report" in source_workbook.sheetnames:
        # Get the source sheet
        source_sheet = source_workbook["Consolidated Test Report"]

        # Create a new sheet in the consolidated workbook with the renamed sheet
        destination_sheet = consolidated_workbook.create_sheet(title=month_year.strftime('%B%Y'))

        # Copy the cell values and formatting from source sheet to destination sheet
        for row in source_sheet.iter_rows(values_only=True):
            destination_sheet.append(row)

# Remove the default sheet created in the consolidated workbook
consolidated_workbook.remove(consolidated_workbook.active)

# Save the consolidated workbook with the copied and renamed sheets
consolidated_workbook.save(destination_file_path)

print('Sheets copied and renamed successfully. Consolidated file created.')
