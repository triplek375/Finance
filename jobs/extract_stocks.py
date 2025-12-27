from bs4 import BeautifulSoup
import requests

# Extract Stocks - LTP, URL, Industry
def run():
    base_url = 'https://www.screener.in'
    stocks_to_url = {}

    market = BeautifulSoup(requests.get(base_url+'/market').text,'html.parser')
    for industry_ele in market.find_all('a', href=lambda href: href and href.startswith('/market/')):
        industry_page = BeautifulSoup(requests.get(base_url + industry_ele['href']).text,'html.parser')
        for industry_ele in industry_page.find_all('a', href=lambda href: href and href.startswith('/company/')):
            stocks_to_url[industry_ele.text] = base_url + industry_ele['href']
        