import os
import xlwings as xw

# Path to the directory containing .xlsx files
directory = '/path/to/directory'

# Name of the VBA macro to run
macro_name = 'MacroName'

# Iterate over each file in the directory
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(directory, filename)

        # Connect to the Excel application
        app = xw.App(visible=False)
        workbook = app.books.open(file_path)

        # Run the VBA macro
        workbook.macro(macro_name).run()

        # Save and close the workbook
        workbook.save()
        workbook.close()

        # Quit the Excel application
        app.quit()
