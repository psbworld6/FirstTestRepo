import os
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Get the directory path where the Excel files are located
directory = '/path/to/excel/files'  # Replace with the actual directory path

# Create a new workbook to store the consolidated sheets
consolidated_workbook = Workbook()

# Iterate over each file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
        # Load the Excel file
        file_path = os.path.join(directory, filename)
        workbook = load_workbook(file_path)

        # Check if the "Consolidated Test File" sheet exists
        if "Consolidated Test File" in workbook.sheetnames:
            # Get the reference to the source sheet
            source_sheet = workbook["Consolidated Test File"]

            # Create a new sheet in the consolidated workbook
            new_sheet = consolidated_workbook.create_sheet(title=filename)

            # Copy the sheet with formatting
            new_sheet.sheet_format = source_sheet.sheet_format
            new_sheet.sheet_properties.tabColor = source_sheet.sheet_properties.tabColor

            for row in source_sheet.iter_rows(min_row=1, max_row=source_sheet.max_row, min_col=1, max_col=source_sheet.max_column):
                for cell in row:
                    new_cell = new_sheet[cell.coordinate]
                    new_cell.data_type = cell.data_type
                    new_cell.value = cell.value
                    if cell.has_style:
                        new_cell.font = cell.font
                        new_cell.fill = cell.fill
                        new_cell.border = cell.border
                        new_cell.alignment = cell.alignment
                        new_cell.number_format = cell.number_format

            # Rename the new sheet to the file name
            new_sheet.title = filename

print("Copying and renaming sheets completed.")

# Save the consolidated workbook to a new file
consolidated_file_path = os.path.join(directory, "consolidated_reports.xlsx")
consolidated_workbook.save(consolidated_file_path)

print(f"Consolidated sheets saved to {consolidated_file_path}.")
