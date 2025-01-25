from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill

def status_check(url):
    return requests.get(url)

def search(soup, body, title):
    return soup.find_all(body, class_=title)

def scrape_page(url, sheet):
    r = status_check(url)
    soup = BeautifulSoup(r.content, "html.parser")

    ratings_list = [ratings.text.strip() for ratings in search(soup, "div", "game-rating-item")]
    titles_list = [elements.text.strip() for elements in search(soup, "div", "title")]
    realese_list = [data.text.strip() for data in search(soup, 'div', 'releases-short')]
    class_list = [classes.text.strip() for classes in search(soup, 'div', 'tags')]
    
    for title, rating, date, tag in zip(titles_list, ratings_list, realese_list, class_list):
        sheet.append([title, date, rating, tag])

    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 20



wb = Workbook()
sheet = wb.active
sheet.append(['Title', 'Release Date', 'Rating', 'Tag'])

base_url = "https://www.playground.ru/games/vr?p="
for page_num in range(1, 21):
    url = f"{base_url}{page_num}"
    print(f"Извлечение данных со страницы: {url}")
    scrape_page(url, sheet)

wb.save("output_data.xlsx")
