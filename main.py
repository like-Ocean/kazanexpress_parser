import xlsxwriter
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup as BS


result_list = []

options = Options()
options.add_argument('--headless')
driver = webdriver.Chrome(options=options)


def get_data():
    try:
        page = 1
        while page <= 100:
            URL = 'https://kazanexpress.ru/category/Odezhda-10014?currentPage=' + str(page)
            driver.get(URL)
            time.sleep(5)
            sours_data = driver.page_source
            soup = BS(sours_data, 'html.parser')
            card = soup.find_all('div', 'ui-card')
            for item in card:
                title = item.find('a', class_='subtitle-item').text
                title = title.replace('\n', '')
                link = "https://kazanexpress.ru" + item.a['href']
                current_price = item.find('span', class_='currency product-card-price slightly medium').text
                old_price = ""
                if item.find('span', class_='currency product-card-old-price regular'):
                    old_price = item.find('span', class_='currency product-card-old-price regular').text
                result_list.append({
                    'name': title.strip(" "),
                    'current_price': current_price.replace('\xa0', ''),
                    'old_price': old_price.replace('\xa0', ''),
                    'link': link
                })
            page += 1
    except Exception as ex:
        print(ex)

    finally:
        driver.close()
        driver.quit()


def create_excel(filename, data):
    if not len(data):
        return None
    with xlsxwriter.Workbook(filename) as workbook:
        try:
            sheets = workbook.add_worksheet()
            header = ['Товар', 'Текущая цена', 'Старая цена', 'Ссылка']

            cell_price_format = workbook.add_format()
            cell_price_format.set_align('center')

            for row, item in enumerate(header):
                sheets.write_string(0, row, item)

            for row, item in enumerate(data, start=1):
                sheets.write_string(row, 0, item['name'])
                sheets.write_string(row, 1, item['current_price'], cell_price_format)
                sheets.write_string(row, 2, item['old_price'], cell_price_format)
                sheets.write_string(row, 3, item['link'])
        except Exception as ex:
            print(ex)


def main():
    get_data()
    create_excel('data.xlsx', result_list)


if __name__ == "__main__":
    main()
