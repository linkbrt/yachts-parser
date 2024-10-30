from openpyxl import load_workbook, Workbook

from bs4 import BeautifulSoup
import requests


CATALOG_URL = "https://yacht-parts.ru/catalog/"


def create_xlsx_file(filename: str):
    workbook = Workbook() 
    sheet = workbook.active

    sheet["A1"] = "Hello"

    workbook.save("main.xlsx")


def parse_catalog():
    resp = requests.get(CATALOG_URL)
    soup = BeautifulSoup(resp.text, "html.parser")

    items = soup.find_all("li", "sect")
    return "\n".join([item.a.get('href') for item in items])


def parse_page_numbers(url):
    resp = requests.get(url)
    soup = BeautifulSoup(resp.text, "html.parser")

    return int([item for item in soup.find("span", "nums")][-2].text)


def parse_item(url):
    resp = requests.get(url)
    soup = BeautifulSoup(resp.text, "html.parser")

    title = soup.find("h1", id="pagetitle").text
    price = soup.find("div", "price").text.strip()
    article = soup.find("div", "article iblock").find("span", "value").text
    description = soup.find("div", "preview_text").text.strip()
    images = soup.find("div", "slides").find_all("img")
    print(title)
    print(price)
    print(article)
    print(description)
    print([image.get("src") for image in images])


if __name__== '__main__':
    
    # resp = requests.get("https://yacht-parts.ru/catalog/exterior/yakornoe/")
    # soup = BeautifulSoup(resp.text, "html.parser")

    # for item in soup.find_all("div", "list_item_wrapp"):
    #     print(item.find("div", "item-title").a.get('href'))

    # print(parse_catalog())
    # parse_page_numbers("https://yacht-parts.ru/catalog/exterior/yakornye_lebedki/")
    parse_item("https://yacht-parts.ru/catalog/comfort/magnitoly-audiosistemy/vodoneproniczaemyj-tyuner-mp6/")
