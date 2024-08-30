import pandas as pd

# Считываем файл Excel
file_path = "file.xlsx"
sheets = pd.read_excel(file_path, sheet_name=None)

target_group = "ИКБО-72-23"

for sheet_name, sheet_data in sheets.items():
    print("Проверка:", sheet_name)

    group_found = False
    for idx, row in sheet_data.iterrows():
        if row.isnull().all():
            continue

        if target_group in row.astype(str).values:
            print(f"Группа {target_group} найдена на листе {sheet_name}, строка {idx}")
            start_idx = int(idx)

            # Предполагаем, что расписание занимает следующие 7 строк
            end_idx = start_idx + 7

            if end_idx <= len(sheet_data):
                schedule = sheet_data.iloc[start_idx:end_idx]

                # Выводим расписание для проверки
                print("Расписание найдено:")
                print(schedule)
            else:
                print("Индексы выходят за пределы данных листа")

            group_found = True
            break

    if not group_found:
        print("Группа не найдена на листе", sheet_name)