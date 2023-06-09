import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Get the directory path where the Excel files are located
directory = '/path/to/excel/files'  # Replace with the actual directory path

# Create a new workbook to store the consolidated sheets
consolidated_workbook = None

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

            # Copy the sheet to the consolidated workbook
            if consolidated_workbook is None:
                consolidated_workbook = Workbook()
                consolidated_sheet = consolidated_workbook.active
                consolidated_sheet.title = filename
            else:
                consolidated_sheet = consolidated_workbook.create_sheet(title=filename)

            # Copy the cell values and formatting
            for row in source_sheet.iter_rows(values_only=True):
                consolidated_sheet.append(row)

            # Apply alignment to the cells in the consolidated sheet
            for row in consolidated_sheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='left')

print("Copying and renaming sheets completed.")

# Save the consolidated workbook to a new file
consolidated_file_path = os.path.join(directory, "consolidated_reports.xlsx")
consolidated_workbook.save(consolidated_file_path)

print(f"Consolidated sheets saved to {consolidated_file_path}.")
