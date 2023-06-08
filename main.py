import collections
import re
from datetime import date
from argparse import ArgumentParser, Namespace
import os
from pandas import read_excel
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

WINERY_FOUNDATION_YEAR = 1920


def get_launch_args() -> Namespace:
    """
    Return the command line arguments for launching the 'Novoe Russkoe Vino' website.

    """

    arg_parser = ArgumentParser(description='Запуск сайта винодельни "Новое русское вино"')
    arg_parser.add_argument('-p', '--path_to_xlsx', help="Путь к файлу с данными о товарах", default="products.xlsx")
    args = arg_parser.parse_args()
    return args


def adjust_year_sign(years: int) -> str:
    """
    Returns the correct grammatical form of the word "year" in Russian, based on the input number of years.
    The input should be an integer.

    """

    if not isinstance(years, int) or years <= 0:
        raise ValueError("Year must be a positive integer!")

    years = str(years)  # Необходимо для работы с регулярными выражениями.
    if re.match(r"^(?!.*11$).*1$", years):
        return "год"
    if re.match(r"^(?!.*1[2-4]$)\d*[2-4]$", years):
        return "года"
    if re.match(r"^\d*[0-9][056789]$", years) or re.match(r"1[1-9]|1[0-9]", years):
        return "лет"


def read_products_xlsx(path: str = "products.xlsx") -> dict[str, list[dict[str, str]]]:
    """
    Reads product information from an Excel file and returns a dictionary where the keys are product categories and
    the values are lists of dictionaries representing the products in that category.

    """

    products_table = read_excel(path, keep_default_na=False)
    products_by_category = collections.defaultdict(list)
    for _, product in products_table.iterrows():
        products_by_category[product["Категория"]].append(product.drop("Категория").to_dict())

    return products_by_category


if __name__ == "__main__":

    arg_vars = get_launch_args()
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    winery_age = date.today().year - WINERY_FOUNDATION_YEAR
    template = env.get_template('template.html')
    rendered_page = template.render(
        winery_age=winery_age,
        year_sign=adjust_year_sign(winery_age),
        product_cards=read_products_xlsx(arg_vars.path_to_xlsx)
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
