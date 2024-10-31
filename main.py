import logging
from datetime import datetime as dt
from multiprocessing.dummy import Pool as ThreadPool

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


BASE_URL = "https://yacht-parts.ru"
CATALOG_URL = "https://yacht-parts.ru/catalog/"
BRANDS_URLS = [
    "https://yacht-parts.ru/info/brands/",
    "https://yacht-parts.ru/info/brands/?PAGEN_1=2",
    "https://yacht-parts.ru/info/brands/?PAGEN_1=3",
]


# Парсинг страницы каталога для получения всех категорий товаров
def parse_catalog() -> dict[str, str]:
    try:
        resp = requests.get(CATALOG_URL)
        soup = BeautifulSoup(resp.text, "html.parser")

        sections = soup.find_all("div", "section_item")
        result = {}
        for item in sections:
            category_name = item.find("li").a.span.text
            result[category_name] = [item.a.get('href') for item in item.find_all("li", "sect")]

    except Exception as e:
        logging.exception(e)
    return result


# Парсинг каждой категории для получения кол-ва страниц пагинатора
def parse_page_numbers(url: str) -> tuple[str, int]:
    try:
        resp = requests.get(BASE_URL + url)
        soup = BeautifulSoup(resp.text, "html.parser")
        try:
            num = int([item for item in soup.find("span", "nums")][-2].text)
        except TypeError:
            num = 1
        logging.info(f"Получил кол-во страниц {url}")

    except Exception as e:
        logging.exception(e)

    return (url, num)


# Парсинг адресов предметов на странице
def parse_page(data: tuple[str, int]) -> list[str]:
    try:
        resp = requests.get(BASE_URL + data[0], params={"?PAGEN_1": data[1]})
        soup = BeautifulSoup(resp.text, "html.parser")
        items = [item.a.get('href') for item in soup.find_all("div", "item-title")]
        logging.info(f"Page parsed: {data[0]}?PAGEN_1={data[1]}")

    except Exception as e:
        logging.exception(e)

    return items


# Парсинг предмета
def parse_item(url: str) -> dict[str, str]:
    result = {}
    try:
        resp = requests.get(BASE_URL + url)
        soup = BeautifulSoup(resp.text, "html.parser")

        title = soup.find("h1", id="pagetitle").text
        price = soup.find("div", "price").text.strip()
        article = soup.find("div", "article iblock").find("span", "value").text
        description = soup.find("div", "preview_text").text.strip()
        images = soup.find("div", "slides").find_all("img")
        logging.info(f"Item parsed: {url}")
        result = {
            "title": title,
            "price": price,
            "article": article,
            "description": description,
            "images": [image.get("src") for image in images],
        }

    except Exception as e:
        logging.exception(e)

    return result


# Парсинг существующих брендов
def parse_brands(url: str) -> list[str]:
    try: 
        resp = requests.get(url)
        soup = BeautifulSoup(resp.text, "html.parser")
        brand_list = list(map(
            lambda item: item["title"], 
            [item.find("img") for item in soup.find("ul", "brands_list")][1:-1:2]
        ))

    except Exception as e:
        logging.exception(e)

    return brand_list


# Сохранение в эксель
def parse_items_and_save(filename: str, catalog: dict[str, str], brands: list[str]):
    workbook = Workbook() 

    for key in catalog:
        urls_with_nums = pool.map(parse_page_numbers, catalog[key])
        workbook_name = key[:31] # Название листа должно содержать < 31 символа
        workbook.create_sheet(workbook_name)
        sheet = workbook[workbook_name]
        all_urls = []
        items = []
        for url, num in urls_with_nums:
            all_urls += pool.map(parse_page, [(url, i) for i in range(num)])
            for page in all_urls:
                items += pool.map(parse_item, [item for item in page])
                break
            break
        sheet.cell(1, 1, "Название")
        sheet.cell(1, 2, "Бренд")
        sheet.cell(1, 3, "Категория")
        sheet.cell(1, 4, "Цена")
        sheet.cell(1, 5, "Артикул")
        sheet.cell(1, 6, "Описание")
        sheet.cell(1, 7, "Изображения")
        for i, item in enumerate(items, 2):
            item_brand = "Нет"
            for brand in brands:
                if brand in item.get("title", ""):
                    item_brand = brand

            sheet.cell(i, 1, item.get("title"))
            sheet.cell(i, 2, item_brand)
            sheet.cell(i, 3, key)
            sheet.cell(i, 4, item.get("price"))
            sheet.cell(i, 5, item.get("article"))
            sheet.cell(i, 6, item.get("description"))
            sheet.cell(i, 7, ', '.join(item.get("images", [])))
        break

    del workbook["Sheet"]
    workbook.save("output/" + filename)


if __name__== '__main__':
    start_time = dt.now()
    pool = ThreadPool(8)
    logging.basicConfig(
        level=logging.INFO,
        format="{asctime} - {levelname} - {message}", 
        filename="logs/app.log",
        style="{",
        filemode="a")

    logging.info("PROGRAM STARTED")

    brands = pool.map(parse_brands, BRANDS_URLS)

    brands = brands[0] + brands[1] + brands[2] # create 1 arr from 3

    catalog = parse_catalog()

    parse_items_and_save(f"main_ {dt.now()}.xlsx", catalog, brands)
    end_time = dt.now()

    logging.info(f"Программа отработала за: {(end_time - start_time).total_seconds()} сек.")
    logging.info("PROGRAM EXITED")
