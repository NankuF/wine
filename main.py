import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

excel_wines = pd.read_excel('wine3.xlsx', na_values=None, keep_default_na=False).to_dict(orient='records')

grouped_products = collections.defaultdict(list)
for product in excel_wines:
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
