import os
from openpyxl import load_workbook
from openpyxl.vba import ProjectModule

# Path to the directory containing .xlsx files
directory = '/path/to/directory'

# Path to the .bas file
bas_file_path = '/path/to/macro.bas'

# Load the .bas file content
with open(bas_file_path, 'r') as bas_file:
    bas_code = bas_file.read()

# Iterate over each file in the directory
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(directory, filename)

        # Load the workbook
        workbook = load_workbook(file_path)

        # Create a new VBA module and add the .bas code
        vba_module = ProjectModule(filename='Module1', content=bas_code)
        workbook.vba_project.vba_modules.append(vba_module)

        # Save the modified workbook
        workbook.save(file_path)
