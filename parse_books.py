"""Демо-парсер для портфолио: собирает каталог учебного магазина books.toscrape.com
(сайт сделан специально для тренировки парсинга) в таблицу Excel.

Собирает 1000 книг: название, категория, цена, рейтинг, остаток, артикул, ссылка.
Проверяет дубли и пустые поля, на втором листе строит сводку формулами.

Запуск:
    pip install requests beautifulsoup4 openpyxl
    python parse_books.py
    python parse_books.py --fast    # без карточек книг: быстрее, но без категории и остатка
"""
import re
import sys
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

START = "https://books.toscrape.com/catalogue/page-1.html"
RATING = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
HEAD_FILL = PatternFill("solid", fgColor="3F3CBB")
HEAD_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")


def get(session, url):
    for attempt in range(3):
        try:
            r = session.get(url, timeout=20)
            r.raise_for_status()
            r.encoding = "utf-8"
            return BeautifulSoup(r.text, "html.parser")
        except requests.RequestException:
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Не удалось загрузить {url}")


def parse_list(session):
    books, url = [], START
    while url:
        soup = get(session, url)
        for card in soup.select("article.product_pod"):
            link = card.select_one("h3 a")
            price = card.select_one("p.price_color").get_text(strip=True)
            rating = card.select_one("p.star-rating")["class"][1]
            books.append({
                "title": link["title"].strip(),
                "url": urljoin(url, link["href"]),
                "price": float(re.sub(r"[^\d.]", "", price)),
                "rating": RATING.get(rating),
            })
        nxt = soup.select_one("li.next a")
        url = urljoin(url, nxt["href"]) if nxt else None
        print(f"\rСтраниц каталога: {len(books) // 20}, книг: {len(books)}", end="")
        time.sleep(0.2)
    print()
    return books


def parse_details(session, books):
    for i, b in enumerate(books, 1):
        soup = get(session, b["url"])
        crumbs = soup.select("ul.breadcrumb li a")
        b["category"] = crumbs[2].get_text(strip=True) if len(crumbs) > 2 else ""
        table = {tr.th.get_text(strip=True): tr.td.get_text(strip=True) for tr in soup.select("table.table tr")}
        b["upc"] = table.get("UPC", "")
        stock = re.search(r"\d+", table.get("Availability", ""))
        b["stock"] = int(stock.group()) if stock else 0
        if i % 50 == 0 or i == len(books):
            print(f"\rКарточек обработано: {i} из {len(books)}", end="")
        time.sleep(0.1)
    print()


def save(books, detailed, path="Книги.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Каталог"
    cols = [("Название", 60), ("Категория", 20), ("Цена, £", 10), ("Рейтинг", 9),
            ("В наличии, шт.", 14), ("Артикул (UPC)", 20), ("Ссылка", 16)]
    ws.append([c for c, _ in cols])
    for i, (_, w) in enumerate(cols, 1):
        cell = ws.cell(row=1, column=i)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        ws.column_dimensions[get_column_letter(i)].width = w
    for b in books:
        ws.append([b["title"], b.get("category", ""), b["price"], b["rating"],
                   b.get("stock", None) if detailed else None, b.get("upc", ""), "открыть"])
        link = ws.cell(row=ws.max_row, column=7)
        link.hyperlink, link.font = b["url"], Font(name="Arial", color="3F3CBB", underline="single")
    for row in ws.iter_rows(min_row=2, max_col=6):
        for c in row:
            c.font = BODY_FONT
    for c in ws["C"][1:]:
        c.number_format = "0.00"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{ws.max_row}"

    last = ws.max_row
    s = wb.create_sheet("Сводка")
    s.column_dimensions["A"].width = 34
    s.column_dimensions["B"].width = 14
    s.append(["Показатель", "Значение"])
    for c in s[1]:
        c.fill, c.font = HEAD_FILL, HEAD_FONT
    s.append(["Книг собрано", f"=COUNTA('Каталог'!A2:A{last})"])
    s.append(["Средняя цена, £", f"=ROUND(AVERAGE('Каталог'!C2:C{last}),2)"])
    s.append(["Самая дорогая, £", f"=MAX('Каталог'!C2:C{last})"])
    for r in range(1, 6):
        s.append([f"Книг с рейтингом {r}", f"=COUNTIF('Каталог'!D2:D{last},{r})"])
    if detailed:
        s.append(["Всего экземпляров в наличии", f"=SUM('Каталог'!E2:E{last})"])
    for row in s.iter_rows(min_row=2):
        for c in row:
            c.font = BODY_FONT
    wb.save(path)
    return path


def main():
    detailed = "--fast" not in sys.argv
    session = requests.Session()
    session.headers["User-Agent"] = "Mozilla/5.0 (portfolio demo parser)"
    books = parse_list(session)
    if detailed:
        parse_details(session, books)

    titles = [b["title"] for b in books]
    dupes = len(titles) - len(set(b["url"] for b in books))
    empty = sum(1 for b in books if not b["title"] or b["price"] is None or b["rating"] is None)
    path = save(books, detailed)
    print(f"Готово: {len(books)} книг, дублей: {dupes}, пустых полей: {empty}. Файл: {path}")


if __name__ == "__main__":
    main()
