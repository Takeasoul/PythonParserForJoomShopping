import asyncio
import aiohttp
import string
import requests
import random
from bs4 import BeautifulSoup as BS
from fake_user_agent import user_agent
import xlsxwriter
import telebot
from urllib.parse import urljoin
import PySimpleGUI as sg

BASE_URL = ""
list_link = []
new_items = []
oldheaders = []
newheaders = []
listofitem = []
data = []
site = ""

headers = ['Название', 'Изображение:']

HEADERS = {"user-agent", user_agent("chrome")}


def dump_to_xlsx(filename, data, headers):
    if not len(data):
        return None

    with xlsxwriter.Workbook(filename) as workbook:
        ws = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        for col, h in enumerate(headers):
            ws.write_string(0, col, h, cell_format=bold)

        for row, item in enumerate(data, start=1):
            for prop_name, prop_value in item.items():
                col = headers.index(prop_name)
                ws.write_string(row, col, prop_value)


def unique_elements(iterable):
    seen = set()
    result = []
    for element in iterable:
        hashed = element
        if isinstance(element, dict):
            hashed = tuple(sorted(element.iteritems()))
        elif isinstance(element, list):
            hashed = tuple(element)
        if hashed not in seen:
            result.append(element)
            seen.add(hashed)
    return result


def get_base_url(getBASE_URL):
    head, sep, tail = getBASE_URL.partition('.ru')
    return head + sep


# Первый цикл вытаскиывающий из главных страниц католога названия товара, цены и ссылки
async def main():
    layout = [
        [sg.Text('Введите ссылку на сайт'), sg.InputText(), sg.Radio('Slavdom', "RADIO1", default=False, key="slav"),
         sg.Radio('Lepninaplast', "RADIO1", default=False, key="lep")],
        [sg.Text('Введите кол-во страниц, которые нужно считать'), sg.InputText()],
        [sg.Output(size=(88, 20))],
        [sg.Submit(), sg.Cancel()]
    ]
    window = sg.Window('ULTRA-parser', layout)
    while True:  # The Event Loop
        event, values = window.read()
        # print(event, values) #debug
        if event in (None, 'Exit', 'Cancel'):
            break
        if event == 'Submit':
            if values['slav'] == True:
                OUT_XLSX_FILENAME = random.choice(string.ascii_letters) + ".xlsx"
                BASE_URL = values[0]
                BASE_COUNT = int(values[1])
                if ("?PAGEN_1=" not in BASE_URL):
                    BASE_URL = BASE_URL + "?PAGEN_1="
                else:
                    BASE_URL.partition("1=")
                for amount in range(1, BASE_COUNT + 1):
                    url = BASE_URL + amount.__str__()
                    async with aiohttp.ClientSession() as session:
                        async with session.get(url) as response:
                            r = await aiohttp.StreamReader.read(response.content)
                        soup = BS(r, "html.parser")
                        items = soup.find_all("div", {"class": "cart"})
                        # Второй цикл, открывает по ссылке каждый товар и вытаскивает оттуда характеристики
                        for item in items:
                            title = item.find("div", {"class": "cart__desc"})
                            price = item.find_all("div", {"class": "cart__price"})
                            pricename = item.find_all("div", {"class": "cart__price-desc"})
                            good = {
                                "Название": title.text.strip(),
                            }
                            for i in range(0, len(price)):
                                oldheaders.insert(len(oldheaders) + 1, pricename[i].text.strip())
                                good[oldheaders[-1]] = price[i].sup.previous_sibling.strip()
                                listofitem.insert(len(listofitem) + 1, price[i].sup.previous_sibling.strip())

                            data.append(good)
                            link = get_base_url(BASE_URL) + item.find("a").get("href")
                            list_link.append(link)
                            async with aiohttp.ClientSession() as session2:
                                async with session2.get(link) as response2:
                                    r2 = await aiohttp.StreamReader.read(response2.content)
                                soup2 = BS(r2, "html.parser")
                                items2 = soup2.find_all("div", {"class": "specifications-table__item"})
                                image = soup2.find('img', {"class": "card-slider__img"}).get('src', '-')
                                image = urljoin(get_base_url(BASE_URL), image)
                                good["Изображение:"] = image
                                for i in range(0, len(items2)):
                                    if i % 2 == 0:
                                        oldheaders.insert(len(oldheaders) + 1, items2[i].text.strip())
                                    else:
                                        good[oldheaders[-1]] = items2[i].text.strip()
                                        listofitem.insert(len(listofitem) + 1, items2[i].text.strip())
                                    # print(f"|{items2[i].text.strip()}")
                                # print(listofitem)
                                listofitem.clear()
                                data.append(good)

                newheaders = headers + unique_elements(oldheaders)
                # print(newheaders)
                dump_to_xlsx(OUT_XLSX_FILENAME, data, newheaders)
                print("Создан файл с названием", OUT_XLSX_FILENAME)
                data.clear()
                newheaders.clear()
                oldheaders.clear()

            elif values['lep'] == True:
                OUT_XLSX_FILENAME = random.choice(string.ascii_letters) + ".xlsx"
                BASE_URL = values[0]
                BASE_COUNT = int(values[1])

                if ("?PAGEN_1=" not in BASE_URL):
                    BASE_URL = BASE_URL + "?PAGEN_1="
                else:
                    BASE_URL.partition("1=")
                for amount in range(1, BASE_COUNT + 1):
                    url = BASE_URL + amount.__str__()
                    async with aiohttp.ClientSession() as session:
                        async with session.get(url) as response:
                            r = await aiohttp.StreamReader.read(response.content)
                        soup = BS(r, "html.parser")
                        items = soup.find_all("div", {"class": "info"})
                        # Второй цикл, открывает по ссылке каждый товар и вытаскивает оттуда характеристики
                        for item in items:
                            title = item.find("div", {"class": "name"})
                            price = item.find("div", {"class": "price"})
                            trs = soup.find('div', class_=('info')).find('table').find_all('tr')
                            good = {
                                "Название": title.text.strip(),
                                "Цена": price.text.strip()
                            }
                            oldheaders.insert(len(oldheaders) + 1, "Цена")

                            for tr in trs[:-1]:
                                trr = tr.find_all('td')
                                oldheaders.insert(len(oldheaders) + 1, trr[0].text.strip())
                                good[oldheaders[-1]] = trr[1].text.strip()
                            data.append(good)
                            image = soup.find('div', {"class": "image"})
                            image = image.find('img').get('src')
                            image = urljoin(get_base_url(BASE_URL), image)
                            good["Изображение:"] = image
                newheaders = headers + unique_elements(oldheaders)
                dump_to_xlsx(OUT_XLSX_FILENAME, data, newheaders)
                print("Создан файл с названием", OUT_XLSX_FILENAME)
                data.clear()
                newheaders.clear()
                oldheaders.clear()


if __name__ == '__main__':
    loop = asyncio.new_event_loop()
    loop.run_until_complete(main())
