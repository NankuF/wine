import collections
import datetime
import os
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    load_dotenv()
    path_to_excel = os.getenv('PATH_TO_EXCEL')

    try:
        products = pd.read_excel(path_to_excel, na_values=None, keep_default_na=False).to_dict(orient='records')
    except ValueError:
        raise ValueError('Файл должен быть в формате Excel')
    grouped_products = collections.defaultdict(list)
    for product in products:
        grouped_products[product['Категория']].append(product)

    company_foundation_year = 1920
    company_age = datetime.datetime.now().year - company_foundation_year

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(goods=grouped_products, company_age=company_age)
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
