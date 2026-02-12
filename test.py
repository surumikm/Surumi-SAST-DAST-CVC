import os
import sys
import xlrd
from xlutils.copy import copy
import xlwt

directory = sys.argv[1]

XML_REPORT = sys.argv[2]
word = os.path.splitext(XML_REPORT)[0]

print("directory:", directory)
print("XML_REPORT:", XML_REPORT)
print("word:", word)

def find_xls_files_with_word(directory, word):
    xls_files = []
    for file in os.listdir(directory):
        if file.endswith('.xls') and word in file:
            xls_files.append(os.path.join(directory, file))
    return xls_files

directory = sys.argv[1]
word = os.path.splitext(XML_REPORT)[0]

xls_files = find_xls_files_with_word(directory, word)
print("XLS files with '{}' in their name:".format(word))

xls_file_dict = {}

for i, xls_file in enumerate(xls_files, start=0):
    print(xls_file)
    xls_file_dict[f'xls_file_{i}'] = xls_file

# Accessing files from the dictionary
for key, value in xls_file_dict.items():
    print(f"{key}: {value}")

# Full path to the existing workbook
#xls_file_path = r'C:\Users\1026661\Downloads\Agreement_SAST_Developer_Report_11-06-2024_0.xls'

# Assuming xls_file_dict contains the file paths as described earlier
for key, xls_file in xls_file_dict.items():
    # Open the existing workbook
    workbook = xlrd.open_workbook(xls_file, formatting_info=True)

    # Create a new workbook and copy contents and formatting from the existing workbook
    new_workbook = copy(workbook)

    # Add a new sheet to the new workbook
    sensitivity_sheet = new_workbook.add_sheet('Document Classification')
    summary_sheet = new_workbook.add_sheet('Summary')

    # Merge cells B3 to F3 and write "TCS Confidential" in the merged cell
    sensitivity_sheet.write_merge(2, 2, 1, 5, 'TCS Confidential', xlwt.easyxf('font: name Calibri, bold on, height 400; align: horiz center, vert center'))

    # Save the modified workbook
    new_workbook.save(xls_file)
