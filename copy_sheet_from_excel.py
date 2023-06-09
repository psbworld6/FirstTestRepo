import os
from openpyxl import load_workbook

# Get the directory path where the Excel files are located
directory = '/path/to/excel/files'  # Replace with the actual directory path

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

            # Create a new sheet with the file name
            new_sheet_name = os.path.splitext(filename)[0]
            new_sheet = workbook.create_sheet(title=new_sheet_name)

            # Copy the contents from the source sheet to the new sheet
            for row in source_sheet.iter_rows(values_only=True):
                new_sheet.append(row)

        # Save the modified workbook back to the file
        workbook.save(file_path)

        print(f"Sheet 'Consolidated Test File' copied and renamed in {filename}")

print("All files processed successfully.")
