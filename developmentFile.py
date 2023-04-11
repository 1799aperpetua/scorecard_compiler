from openpyxl import Workbook, load_workbook
import os

# Load the destination workbook where we want all of the sheets to end up
destination = 'source_workbook.xlsx'
dest_wb = load_workbook('source_workbook.xlsx')

# Loop through each .xlsx file in our scorecard_package folder, adding each spreadsheet to the destination workbook
directory = '/scorecard_package'
print("Looping through files in: ", directory, '\n', os.listdir(directory))

for file in os.listdir(directory):
    if file.endswith('.xlsx'):
        cur_wb = load_workbook(f'{os.path.join(directory, file)}')
        ws = cur_wb['scorecard']
        dest_wb.copy_worksheet(ws)
        print('Scorecard submitted for: ', ws['a1'].value)

print(f'All scorecards in {directory} have been added to {destination}!\n')


