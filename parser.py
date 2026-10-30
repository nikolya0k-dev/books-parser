"""
Парсер каталога книг с сайта https://books.toscrape.com

Собирает по каждой книге: название, цену, рейтинг (1–5), наличие и ссылку,
проходит по всем страницам каталога и сохраняет результат в books.xlsx.
"""

import random
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Настройки
# ---------------------------------------------------------------------------

# Первая страница каталога. Дальше парсер сам найдёт ссылку «next».
START_URL = "https://books.toscrape.com/catalogue/page-1.html"

# Имя итогового Excel-файла
OUTPUT_FILE = "books.xlsx"

# Сколько раз пробовать загрузить страницу, прежде чем сдаться
MAX_RETRIES = 3

# Пауза между запросами (в секундах) — случайное число из этого диапазона,
# чтобы не нагружать сайт
DELAY_RANGE = (0.5, 1.0)

# Сколько секунд ждать ответа сервера, прежде чем считать запрос неудачным
TIMEOUT = 10

# На сайте рейтинг записан словом в CSS-классе: <p class="star-rating Three">.
# Этот словарь переводит слово в число.
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

# Заголовок User-Agent: представляемся браузером, а не «python-requests»
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )
}


def fetch_page(session, url):
    """
    Загружает страницу и возвращает её HTML-текст.
    Если загрузка не удалась — повторяет до MAX_RETRIES раз.
    Если все попытки неудачны — возвращает None.
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(url, headers=HEADERS, timeout=TIMEOUT)
            # raise_for_status() выбросит исключение, если код ответа 4xx или 5xx
            response.raise_for_status()
            # Сайт отдаёт UTF-8, но не всегда сообщает об этом в заголовках.
            # Без этой строки символ «£» превращается в «Â£».
            response.encoding = "utf-8"
            return response.text
        except requests.RequestException as error:
            print(f"  ! Ошибка при загрузке (попытка {attempt}/{MAX_RETRIES}): {error}")
            if attempt < MAX_RETRIES:
                # Перед повтором ждём чуть дольше с каждой попыткой
                time.sleep(attempt * 2)
    return None


def parse_book(card, page_url):
    """
    Достаёт данные одной книги из её «карточки» на странице каталога.
    card — это HTML-блок <article class="product_pod">.
    """
    # Полное название лежит в атрибуте title ссылки внутри <h3>
    # (в самом тексте ссылки длинные названия обрезаны «...»)
    link_tag = card.h3.a
    title = link_tag["title"]

    # Ссылка на странице относительная (например, "a-light-in-the-attic_1000/index.html"),
    # urljoin превращает её в полный адрес относительно текущей страницы
    url = urljoin(page_url, link_tag["href"])

    # Цена в виде "£51.77": убираем знак фунта и превращаем в число
    price_text = card.select_one("p.price_color").get_text(strip=True)
    price = float(price_text.replace("£", ""))

    # Рейтинг: у тега классы ["star-rating", "Three"] — берём второе слово
    rating_classes = card.select_one("p.star-rating")["class"]
    rating = next((RATING_WORDS[c] for c in rating_classes if c in RATING_WORDS), None)

    # Наличие: текст вида "In stock" или "Out of stock"
    availability_text = card.select_one("p.availability").get_text(strip=True)
    in_stock = "В наличии" if availability_text.lower() == "in stock" else "Нет в наличии"

    return {
        "title": title,
        "price": price,
        "rating": rating,
        "availability": in_stock,
        "url": url,
    }


def scrape_all_books():
    """
    Обходит все страницы каталога по ссылке «next» и возвращает список книг.
    """
    books = []
    url = START_URL
    page_number = 1
    failed_in_row = 0  # сколько страниц подряд не удалось загрузить

    # Session переиспользует одно соединение для всех запросов — так быстрее
    with requests.Session() as session:
        while url:
            print(f"Парсинг страницы {page_number}: {url}")
            html = fetch_page(session, url)

            if html is None:
                failed_in_row += 1
                print(f"  ! Страница {page_number} пропущена")
                # Две неудачи подряд — скорее всего, страницы закончились
                # или сайт недоступен. Останавливаемся.
                if failed_in_row >= 2:
                    print("  ! Две страницы подряд не загрузились — останавливаемся")
                    break
                # Ссылку «next» с несработавшей страницы не взять, поэтому
                # угадываем адрес следующей по шаблону page-N.html
                page_number += 1
                url = urljoin(START_URL, f"page-{page_number}.html")
                time.sleep(random.uniform(*DELAY_RANGE))
                continue

            failed_in_row = 0
            soup = BeautifulSoup(html, "html.parser")

            # Каждая книга на странице — это <article class="product_pod">
            cards = soup.select("article.product_pod")
            for card in cards:
                books.append(parse_book(card, url))

            print(f"  + найдено книг на странице: {len(cards)}, всего собрано: {len(books)}")

            # Ищем кнопку «next». На последней странице её нет — цикл закончится.
            next_link = soup.select_one("li.next a")
            if next_link:
                url = urljoin(url, next_link["href"])
                page_number += 1
                # Пауза перед следующим запросом
                time.sleep(random.uniform(*DELAY_RANGE))
            else:
                url = None

    return books


def save_to_excel(books, filename):
    """
    Сохраняет список книг в Excel: жирные заголовки, автоширина колонок,
    закреплённая первая строка и фильтр.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Книги"

    # Пары «ключ в словаре книги» -> «заголовок колонки в Excel»
    columns = [
        ("title", "Название"),
        ("price", "Цена (£)"),
        ("rating", "Рейтинг"),
        ("availability", "Наличие"),
        ("url", "Ссылка"),
    ]

    # Строка заголовков
    ws.append([header for _, header in columns])
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Строки с данными
    for book in books:
        ws.append([book[key] for key, _ in columns])

    # Формат цены: два знака после запятой
    for cell in ws["B"][1:]:
        cell.number_format = "0.00"

    # Автоширина: ширина колонки = длина самого длинного значения + небольшой запас.
    # Ограничиваем максимум, чтобы очень длинные названия не растягивали лист.
    for col_index, (key, header) in enumerate(columns, start=1):
        values = [header] + [str(book[key]) for book in books]
        max_length = max(len(v) for v in values)
        ws.column_dimensions[get_column_letter(col_index)].width = min(max_length + 2, 80)

    # Закрепляем строку заголовков (она не уезжает при прокрутке)
    ws.freeze_panes = "A2"
    # Включаем фильтр-стрелочки в заголовках
    ws.auto_filter.ref = ws.dimensions

    wb.save(filename)


def main():
    print("Старт парсинга books.toscrape.com\n")
    books = scrape_all_books()
    print(f"\nГотово! Всего собрано книг: {len(books)}")

    save_to_excel(books, OUTPUT_FILE)
    print(f"Результат сохранён в файл {OUTPUT_FILE}")


# Этот блок выполняется только при запуске файла напрямую (python parser.py),
# но не при импорте его в другой скрипт
if __name__ == "__main__":
    main()
