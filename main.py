import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

excel_wines = pd.read_excel('wine3.xlsx', na_values=None, keep_default_na=False).to_dict(orient='records')

categories = []
for product_dict in excel_wines:
    if product_dict.get('Категория') not in categories:
        categories.append(product_dict.get('Категория'))

goods = {category: [] for category in categories}
for category in categories:
    for product_dict in excel_wines:
        if category == product_dict.get('Категория'):
            goods[category].append(product_dict)

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

date_create_company = datetime.date(year=1920, month=1, day=1)
delta = (datetime.date.today() - date_create_company).days // 365

rendered_page = template.render(goods=goods, years=delta)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
