import math
from bs4 import BeautifulSoup
import os
import shutil
import json
import time
import openpyxl
import asyncio
import aiohttp

# По умолчанию данные сохраняются в папке example !!!

products_data = []
search_item = ""
chosen_category_name = ""
start_time = 0
id = 1
async def get_page_data(session, lang, filter, page):
    headers = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:138.0) Gecko/20100101 Firefox/138.0"
    }
    url = f"https://www.rdveikals.lv/search/{lang}/word/{search_item}/page/{page}/{filter}"

    async with session.get(url=url, headers=headers) as response:
        response_text = await response.text()
        soap = BeautifulSoup(response_text, "lxml")

        items = soap.find_all("li", class_="col col--xs-4 product js-product js-touch-hover")

        for item in items:
            global id
            id += 1
            products_data.append({
                "id": id,
                "name": item.get("data-prod-name"),
                "price": float(item.get("data-prod-price")),
                "link": f"https://www.rdveikals.lv/{item.find('a', class_='overlay').get('href')}"
            })

        print(f"Обработана страница: {page}")

async def gather_data():
    path = "data"
    if os.path.exists(path):
        shutil.rmtree(path)

    headers = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:138.0) Gecko/20100101 Firefox/138.0"
    }
    while (True):
        answer_language = input("Введите желаемый язык:\n 1 - русский\n 2 - латышский\n")
        if answer_language.isnumeric():
            if int(answer_language) == 1:
                language = "ru"
                break
            elif int(answer_language) == 2:
                language = "lv"
                break

    global search_item
    search_item = input("Введите товар, который хотите найти: ")
    url = f"https://www.rdveikals.lv/search/{language}/word/{search_item}/page/1/"

    async with aiohttp.ClientSession() as session:
        response = await session.get(url=url, headers=headers)

        if not os.path.exists("data"):
            os.mkdir('data')

        response_text = await response.text()
        soup = BeautifulSoup(response_text, "lxml")

        filter = ""
        try:
            categories = soup.find("ul", class_="carousel__list").find_all("div", class_="shopping_cart_category__title")
            i = 1
            for category in categories:
                category_name = category.find("a").text
                print(f"{i} - Найдена категория: {category_name}")
                i += 1

            try:
                chosen_category = int(input("Выберите категорию, которая вас интересует:"))
                global chosen_category_name
                chosen_category_name = categories[chosen_category-1].find("a").text
                category_url = categories[chosen_category - 1].find("a").get("href")
                parts = category_url.split("/")
                if "filters" in parts:
                    filters_index = parts.index("filters")
                    filters_value = parts[filters_index + 1]
                    filter = f"filters/{filters_value}/"
            except Exception as e:
                print("Произошла ошибка, категория не была выбрана")
        except:
            print("Такого товара нет")
            return None

        response = await session.get(url=url+filter, headers=headers)

        response_text = await response.text()
        soup = BeautifulSoup(response_text, "lxml")

        pages_count = 1
        try:
            list = soup.find("div", class_="block").find("div", class_="group").find_next("div").find_all("a")
            if len(list) == 1:
                pages_count = 1 + 1
            elif len(list) > 1:
                pages_count = int(list[-1].text)
        except:
            print("Найдена только одна страница")
        tasks = []

        global start_time
        start_time = time.time()
        for page in range(1, pages_count+1):
            task = asyncio.create_task(get_page_data(session, language, filter, page))
            tasks.append(task)

        await asyncio.gather(*tasks)

def make_excel_document(path):
    table_time_start = time.time()
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
    for product in products_data:
        sheet[row][0].value = product["name"]
        sheet[row][0].hyperlink = product["link"]
        sheet[row][0].style = "Hyperlink"
        sheet[row][1].value = f"{product['price']} €"
        print(f"Готово на: {math.ceil((row-1)*100/len(data))}%")
        row += 1

    book.save(f"example/items_{search_item}{'_'+chosen_category_name if chosen_category_name else ''}.xlsx")
    book.close()

    finish_table_time = round(time.time() - table_time_start, 2)
    print(f"Таблица была создана за: {finish_table_time} с")

def main():
    asyncio.run(gather_data())
    finish_time = time.time() - start_time
    print(f"Сбор данных был завершен за: {round(finish_time, 2)} с")
    with open(f"example/items_list_{search_item}{'_'+chosen_category_name if chosen_category_name else ''}.json", "w", encoding="utf-8") as file:
        json.dump(products_data, file, indent=4, ensure_ascii=False)

    make_document = input("Все данные получены, хотите ли вы их объединить в Excel таблице? y/n\n")
    if (make_document == "y"):
        path = f"example/items_list_{search_item}{'_'+chosen_category_name if chosen_category_name else ''}.json"
        make_excel_document(path)
        print("Документ успешно сохранен!")
    else:
        print("Приятного дня")


if __name__ == "__main__":
    main()