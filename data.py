import openpyxl
import json
from datetime import datetime

book = openpyxl.load_workbook("file.xlsx")
sheet = book.active

data = []

for row in sheet.iter_rows(values_only=True):
    row_data = {}
    for index, cell_value in enumerate(row):
        row_data[f'Column{index + 1}'] = cell_value
    data.append(row_data)
    #print(row_data)

json_data = json.dumps(data, indent=4)

with open("shedule.json", "w") as json_file:
    json_file.write(json_data)

print(datetime.today().weekday())