import os
import xlwings as xw

# Path to the directory containing .xlsx files
directory = '/path/to/directory'

# Path to the .bas file
bas_file_path = '/path/to/macro.bas'

# Iterate over each file in the directory
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(directory, filename)

        # Connect to the Excel application
        app = xw.App(visible=False)
        workbook = app.books.open(file_path)

        # Load the .bas file content
        with open(bas_file_path, 'r') as bas_file:
            bas_code = bas_file.read()

        # Run the VBA code
        workbook.api.Application.Run(bas_code)

        # Save and close the workbook
        workbook.save()
        workbook.close()

        # Quit the Excel application
        app.quit()
