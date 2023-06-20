from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter

def apply_conditional_formatting_to_all_sheets(file_path):
    # Load the workbook
    workbook = load_workbook(filename=file_path)
    
    # Set the conditional formatting properties
    fill = PatternFill(start_color="FFFFD0", end_color="FFFFD0")
    font = Font(color="FF0000")
    
    # Iterate through each sheet in the workbook
    for sheet in workbook.sheetnames:
        # Get the sheet object
        ws = workbook[sheet]
        
        # Set the range for conditional formatting (F:P)
        start_column = "F"
        end_column = "P"
        start_row = 1
        end_row = ws.max_row
        
        # Apply conditional formatting to the range
        for col in range(ord(start_column), ord(end_column) + 1):
            column_letter = get_column_letter(col)
            column_range = f"{column_letter}{start_row}:{column_letter}{end_row}"
            
            rule = Rule(type="uniqueValues")
            rule.dxf = DifferentialStyle(fill=fill, font=font)
            ws.conditional_formatting.add(column_range, rule)
    
    # Save the workbook
    workbook.save(filename=file_path)
    workbook.close()
    
    print("Conditional formatting applied to all sheets.")

# Specify the file path of the workbook
file_path = "C:/Your/Workbook/Path.xlsx"

# Call the function to apply conditional formatting to all sheets
apply_conditional_formatting_to_all_sheets(file_path)
