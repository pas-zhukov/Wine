import collections
import re
from datetime import date
from pandas import read_excel
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

WINERY_FOUNDATION_DATE = date(1920, 1, 1)


def adjust_year_sign(years: int) -> str:
    """
    Returns the correct grammatical form of the word "year" in Russian, based on the input number of years.
    The input should be an integer.

    """

    if not isinstance(years, int) or years <= 0:
        raise ValueError("Year must be a positive integer!")
    years = str(years)

    if re.match(r"^(?!.*11$).*1$", years):
        return "год"
    if re.match(r"^(?!.*1[2-4]$)\d*[2-4]$", years):
        return "года"
    if re.match(r"^\d*[0-9][056789]$", years) or re.match(r"1[1-9]|1[0-9]", years):
        return "лет"


def read_products_info(path: str = "wine3.xlsx") -> dict[str, list[dict[str, str]]]:
    """
    Reads product information from an Excel file and returns a dictionary where the keys are product categories and
    the values are lists of dictionaries representing the products in that category.

    """

    data = read_excel(path, keep_default_na=False)
    dict_data = collections.defaultdict(list)
    for i, row in data.iterrows():
        dict_data[row["Категория"]].append(row.drop("Категория").to_dict())

    return dict_data


if __name__ == "__main__":
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    winery_age = date.today().year - WINERY_FOUNDATION_DATE.year
    template = env.get_template('template.html')
    rendered_page = template.render(
        winery_age=winery_age,
        year_sign=adjust_year_sign(winery_age),
        products_data=read_products_info()
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
