import os
import configparser
import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    import openpyxl
except ImportError:
    print("openpyxl is not installed. Installing...")
    install("openpyxl")

def create_School_ini_from_excel(file_path, output_file):
    # Load the Excel file
    wb = openpyxl.load_workbook(file_path)
    sheet = wb['Sheet1']

    # Create a new INI file
    config = configparser.ConfigParser()

    # Iterate over rows starting from the second row
    for row in sheet.iter_rows(min_row=2, values_only=True):
        column_a = str(row[0]).strip()  # Get value from column A
        section_title = column_a

        # Skip rows where value is 'None'
        if section_title == 'None':
            continue

        # Skip rows where value is empty
        if not section_title:
            continue

        # Create a new section in the INI file
        config[section_title] = {}

        # Iterate over columns starting from the second column
        for col_index, cell_value in enumerate(row[1:], start=1):
            config[section_title]['System Name'] = str(cell_value).strip()

    # Save the INI file
    with open(output_file, 'w') as config_file:
        config.write(config_file)

    print(f'Successfully created {output_file}')

script_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
# Create School ini's

output_ini_file = os.path.join(script_directory, 'output.ini')
create_School_ini_from_excel(os.path.join(script_directory, 'rename.xlsx'), output_ini_file)