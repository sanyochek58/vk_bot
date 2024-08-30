import requests
from bs4 import BeautifulSoup

# Получаем страницу
page = requests.get("https://www.mirea.ru/schedule/")
soup = BeautifulSoup(page.text, "html.parser")

# Ищем нужный блок с расписанием
schedule_div = soup.find("div", {"class": "schedule"})
if schedule_div:
    print("Schedule div found")

    # Поиск института информационных технологий
    iit_div = schedule_div.find(string="Институт информационных технологий")
    if iit_div:
        print("Институт информационных технологий found")

        # Поднимаемся к родительским элементам
        parent_div = iit_div.find_parent("div").find_parent("div")
        if parent_div:
            print("Parent div found")

            # Ищем все карточки с расписанием
            result = parent_div.find_all("a",
                                         class_="uk-link-toggle")
            print(result)
            for x in result:
                link = x.get('href')
                # Проверяем, относится ли карточка к 1 курсу и содержит ли ссылку
                if "1 курс" in x.text and link:
                    file_url = link
                    print(f"Скачиваем файл по ссылке: {file_url}")

                    # Скачиваем и сохраняем файл
                    with open("file.xlsx", "wb") as f:
                        resp = requests.get(file_url)
                        f.write(resp.content)
                    print("Файл сохранён")
                    break
        else:
            print("Parent div not found")
    else:
        print("Институт информационных технологий not found")
else:
    print("Schedule div not found")