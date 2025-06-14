import math
import requests
from bs4 import BeautifulSoup
import os
import shutil
import json
import time
import openpyxl

def get_all_pages():
    path = "data"
    if os.path.exists(path):
        shutil.rmtree(path)

    headers = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:138.0) Gecko/20100101 Firefox/138.0"
    }
    while(True):
        language = input("Введите желаемый язык:\n 1 - русский\n 2 - латышский\n")
        if language.isnumeric():
            if int(language)  == 1:
                lang = "ru"
                break
            elif int(language)  == 2:
                lang = "lv"
                break


    search_item = input("Введите товар, который хотите найти: ")
    url = f"https://www.rdveikals.lv/search/{lang}/word/{search_item}/page/1/"
    r = requests.get(url=url, headers=headers)

    if not os.path.exists("data"):
        os.mkdir('data')

    soup = BeautifulSoup(r.content, "lxml")

    categories_block = soup.find("ul", class_="carousel__list")
    if categories_block:
        categories = categories_block.find_all("div", class_="shopping_cart_category__title")

        i = 1
        for category in categories:
            category_name = category.find("a").text
            print(f"{i} - Найдена категория: {category_name}")
            i += 1

        filter = ""
        try:
            chosen_category = int(input("Выберите категорию, которая вас интересует:"))
            category_url = categories[chosen_category - 1].find("a").get("href")
            parts = category_url.split("/")
            filter = ""
            if "filters" in parts:
                filters_index = parts.index("filters")
                filters_value = parts[filters_index + 1]
                filter = f"filters/{filters_value}/"
        except Exception as e:
            print("Произошла ошибка, категория не была выбрана")
    else:
        print("Такого товара нет")
        return None

    r = requests.get(url=url+filter, headers=headers)
    time.sleep(2)

    soup = BeautifulSoup(r.text, "lxml")

    div_block = soup.find("div", class_="block")

    pages_count = 1
    if div_block:
        group = div_block.find("div", class_="group")
        if group:
            list = group.find_next("div").find_all("a")
            if len(list) == 1:
                pages_count = 1+1
            elif len(list) > 1:
                pages_count = int(list[-1].text)
    else:
        print("Ничего не было найдено!")
        return None

    for i in range(1, pages_count+1):
        with open(f"data/page_{i}.html", "w", encoding="utf-8") as file:
            req = requests.get(f"https://www.rdveikals.lv/search/{lang}/word/{search_item}/page/{i}/{filter}")
            file.write(req.text)
        print(f"Готово на: {math.ceil((i * 100) / pages_count)}%")
    return pages_count

def data_scrap(pages_count): # Будем передавать сколько страниц хотим изучить
    count = 0
    items_list = []
    for i in range(1, pages_count + 1):
        with open(f"data/page_{i}.html", "r", encoding="utf-8") as file:
            src = file.read()
        soap = BeautifulSoup(src, "lxml")
        items = soap.find_all("li", class_="col col--xs-4 product js-product js-touch-hover")

        for item in items:
            item_desc = {}
            item_desc["name"] = item.get("data-prod-name")
            item_desc["price"] = float(item.get("data-prod-price"))
            item_desc["link"] = f"https://www.rdveikals.lv/{item.find('a', class_='overlay').get('href')}"

            items_list.append(item_desc.copy())

        count += len(items)
        print(f"Готово на: {math.ceil((i * 100) / pages_count)}%")

    with open("example/items_list.json", "w", encoding="utf-8") as file:
        json.dump(items_list, file, indent=4, ensure_ascii=False)

    print(f"Всего просканировано: {count} товаров")
def make_excel_document(path):
    book = openpyxl.Workbook()

    sheet = book.active

    for i in range(1, 1001):
        for j in range(1, 1001): # Очистка
            sheet.cell(row=i, column=j).value = ""

    with open(path, "r", encoding="utf-8") as file:
        data = json.load(file)

    sheet["A1"] = "PRODUCT"
    sheet["B1"] = "PRICE (€)"

    row = 2
    for product in data:
        sheet[row][0].value = product["name"]
        sheet[row][0].hyperlink = product["link"]
        sheet[row][0].style = "Hyperlink"
        sheet[row][1].value = f"{product['price']} €"
        print(f"Готово на: {math.ceil((row-1)*100/len(data))}%")
        row += 1

    book.save("items.xlsx")
    book.close()

def main():
    pages_quantity = get_all_pages()
    if pages_quantity:
        while True:
            pages_count = int(input(f"Сколько страниц вы хотите просмотреть? Всего доступно: {pages_quantity}\n"))
            if pages_count > pages_quantity:
                print("Вводите валидное значение")
            else:
                break
        data_scrap(pages_count)

        make_document = input("Хочешь ли ты составить все данные в Excel документе? y/n\n")

        if (make_document == "y"):
            path = "example/items_list.json"
            make_excel_document(path)
            print("Документ успешно сохранен!")
        else:
            print("Приятного дня")
    else:
        print("Вводите валидное значение")

if __name__ == "__main__":
    main()