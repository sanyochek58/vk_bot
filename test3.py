import openpyxl

# Указать путь к файлу
file_path = 'file.xlsx'

# Открыть файл и прочитать его
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

# Указать столбцы с интересующими данными
group_column = 0
name_column = 1
lesson_column = 4

# Искать группу ИКБО-72-23
for row in sheet.iter_rows(min_row=2):
    group = row[group_column].value
    if group == 'ИКБО-72-23':
        group_name = row[name_column].value
        lesson = row[lesson_column].value
        print(f'Группа: {group_name}, Название предмета: {lesson}')
        break