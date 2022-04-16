import collections
import datetime
import os
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape

load_dotenv()
path_to_excel = os.getenv('PATH_TO_EXCEL')

try:
    products = pd.read_excel(path_to_excel, na_values=None, keep_default_na=False).to_dict(orient='records')
except ValueError:
    raise ValueError('Файл должен быть в формате Excel')

grouped_products = collections.defaultdict(list)
for product in products:
    grouped_products[product['Категория']].append(product)

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

date_create_company = datetime.date(year=1920, month=1, day=1)
delta = (datetime.date.today() - date_create_company).days // 365

rendered_page = template.render(goods=grouped_products, years=delta)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
