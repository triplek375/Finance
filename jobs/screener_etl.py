import time
from bs4 import BeautifulSoup
import requests
from datetime import datetime
import pytz
import gspread

class Screener():
  def __init__(self):
    self.base_url = 'https://www.screener.in'
    self.industry_to_url = {}
    self.stock_to_data = {}
    self.epoch_time = time.time()
    self.stock = None

  def get_industry_to_url(self):
    soup = BeautifulSoup(requests.get(self.base_url+'/market').text,'html.parser')
    for ele in soup.find_all('a', href=lambda href: href and href.startswith('/market/')):
      self.industry_to_url[ele.text] = self.base_url + ele['href']

  def get_price_and_ratios(self):
    for ele in self.stock_to_data[self.stock]['Soup'].find_all('span',{'class':'name'}):
      if ele.text.strip() == 'Current Price':
        self.stock_to_data[self.stock]['LTP'] = float(ele.find_next('span',{'class':'number'}).text.replace(',',''))
      elif ele.text.strip() == 'High / Low':
        self.stock_to_data[self.stock]['52W H'] = float(ele.find_all_next('span',{'class':'number'})[0].text.replace(',',''))
        self.stock_to_data[self.stock]['52W L'] = float(ele.find_all_next('span',{'class':'number'})[1].text.replace(',',''))
      elif ele.text.strip() == 'Stock P/E':
        self.stock_to_data[self.stock]['PE'] = float(ele.find_next('span',{'class':'number'}).text.replace(',','')) if ele.find_next('span',{'class':'number'}).text.replace(',','')!='' else 0

  def get_quarterly_results(self):
    section = self.stock_to_data[self.stock]['Soup'].find('section',{'id':'quarters'})
    if section:
      self.stock_to_data[self.stock]['Reported_Upto'] = section.find('table',{'class':'data-table responsive-text-nowrap'}).find('thead').find_all('th')[-1].text.strip()
      if section.find('span',{'class':'badge'}):
        self.stock_to_data[self.stock]['Upcoming_Date'] = section.find('span',{'class':'badge'}).find('strong').text

  def get_profit_loss(self):
    section = self.stock_to_data[self.stock]['Soup'].find('section',{'id':'profit-loss'})
    if section:
      for row in section.find('table',{'class':'data-table responsive-text-nowrap'}).find_all('tr'):
        if not self.stock_to_data[self.stock]['Reported_Upto']:
          self.stock_to_data[self.stock]['Reported_Upto'] = row.find_all('th')[-1].text.strip()
        button = row.find('button')
        if button:
          txt = button['onclick'].strip()
          type = txt[txt.find("('")+2:txt.find("',")]
        else:
          type = row.find('td').text.strip() if row.find('td') else None
        data = [float(col.text.strip().replace(',','').replace('%',''))
          for col in row.find_all('td')[1:] if col.text.strip()]
        if type in ['Sales','Revenue'] and len(data)>1:
          self.stock_to_data[self.stock]['Sales_TTM'] = data[-1]
        elif type in ['Net Profit'] and len(data)>1:
          self.stock_to_data[self.stock]['Profit_TTM'] = data[-1]
      for tbl in section.find_all('table',{'class':'ranges-table'}):
        th = tbl.find('th').text
        rows = tbl.find_all('td')
        if th=='Compounded Sales Growth':
          self.stock_to_data[self.stock]['Sales_Growth_10Y'] = rows[1].text.strip().replace('%','')
          self.stock_to_data[self.stock]['Sales_Growth_5Y'] = rows[3].text.strip().replace('%','')
          self.stock_to_data[self.stock]['Sales_Growth_3Y'] = rows[5].text.strip().replace('%','')
          self.stock_to_data[self.stock]['Sales_Growth_TTM'] = rows[7].text.strip().replace('%','')
        elif th=='Compounded Profit Growth':
          self.stock_to_data[self.stock]['Profit_Growth_10Y'] = rows[1].text.strip().replace('%','')
          self.stock_to_data[self.stock]['Profit_Growth_5Y'] = rows[3].text.strip().replace('%','')
          self.stock_to_data[self.stock]['Profit_Growth_3Y'] = rows[5].text.strip().replace('%','')
          self.stock_to_data[self.stock]['Profit_Growth_TTM'] = rows[7].text.strip().replace('%','')

  def get_margin_data(self):
    if self.stock_to_data[self.stock]['Sales_TTM'] not in [0,None]:
      self.stock_to_data[self.stock]['NPM_TTM'] = round((self.stock_to_data[self.stock]['Profit_TTM']/self.stock_to_data[self.stock]['Sales_TTM']),2)*100

  def get_fii_data(self):
    section = self.stock_to_data[self.stock]['Soup'].find('section',{'id':'shareholding'})
    if section:
      tbl = section.find('table',{'class':'data-table'})
      for row in tbl.find_all('tr'):
        button = row.find('button')
        if button:
          txt = button['onclick'].strip()
          type = txt[txt.find("('")+2:txt.find("',")]
        else:
          type = row.find('td').text.strip() if row.find('td') else None
        data = [float(col.text.strip().replace(',','').replace('%',''))
          for col in row.find_all('td')[1:] if col.text.strip()]
        if type == 'foreign_institutions' and len(data)>7:
          fii_ttm = round(sum(data[-4:])/4,2)
          fii_pttm = round(sum(data[-8:-4])/4,2)
          fii_ttm_pttm = fii_ttm-fii_pttm
          self.stock_to_data[self.stock]['FII_TTM_PTTM'] = fii_ttm_pttm

  def get_stock_to_data(self):
    for i, (industry, url) in enumerate(self.industry_to_url.items(), start=1):
      time.sleep(1)
      stocks_url = {ele.text : self.base_url+ele['href']
        for ele in BeautifulSoup(requests.get(url).text,'html.parser').find_all('a', href=lambda href: href and href.startswith('/company/'))}
      for s, (self.stock, stock_url) in enumerate(stocks_url.items(), start=1):
        time.sleep(0.5)
        soup = BeautifulSoup(requests.get(stock_url).text,'html.parser')
        self.stock_to_data[self.stock] = {'url':stock_url,'Industry':industry,'Soup':soup, 'symbol':stock_url.split('/')[4],
                                          'LTP':None,'52W L':None,'52W H':None,'PE':None,
                                          'Reported_Upto':None,'Upcoming_Date':None,
                                          'Sales_TTM':None,'Profit_TTM':None,'NPM_TTM':None,
                                          'Sales_Growth_10Y':None,'Sales_Growth_5Y':None,'Sales_Growth_3Y':None,'Sales_Growth_TTM':None,
                                          'Profit_Growth_10Y':None,'Profit_Growth_5Y':None,'Profit_Growth_3Y':None,'Profit_Growth_TTM':None,
                                          'FII_TTM_PTTM':None}
        print(f"\r{industry}[{i}/{len(self.industry_to_url)}] {self.stock}[{s}/{len(stocks_url)}] ", end="", flush=True)
        self.get_price_and_ratios()
        self.get_quarterly_results()
        self.get_profit_loss()
        self.get_margin_data()
        self.get_fii_data()
        self.stock_to_data[self.stock]['Soup'] = None
        if not self.stock_to_data[self.stock]['LTP']:
          print(stock_url)

  def update_data_to_sheets(self,gc):
    doc = gc.open('Screener Tracker')
    stocks = doc.worksheet('Stocks')
    stocks.clear()
    stocks.clear_notes(['A','H','S'])
    stocks.clear_basic_filter()
    stocks_cells = []
    stocks_notes = {}
    stocks_cells.append(gspread.Cell(row=1,col=1,value='Stock'))
    timestamp = datetime.fromtimestamp(self.epoch_time,tz=pytz.timezone('Asia/Kolkata')).strftime('%Y-%m-%d %H:%M:%S')
    stocks_notes['A1'] = f'Timestamp : {timestamp}'
    stocks_notes['E1'] = '(LTP-52WL) X 100\n--------------------------\n(52WH-52WL)'
    headers = [
        'Last\nTraded\nPrice',
        '52\nWeek\nLow',
        '52\nWeek\nHigh',
        'Normal\nScore',
        'P/E',
        'Sales\nTTM\n(Cr)',
        'Profit\nTTM\n(Cr)',
        'NPM\nTTM\n%',
        'Sales\nCAGR\n10Y%',
        'Sales\nCAGR\n5Y%',
        'Sales\nCAGR\n3Y%',
        'Sales\nCAGR\nTTM%',
        'Profit\nCAGR\n10Y%',
        'Profit\nCAGR\n5Y%',
        'Profit\nCAGR\n3Y%',
        'Profit\nCAGR\nTTM%',
        'FIIs\nChange\nTTM',
        'Reported Upto/\nUpcoming Date',
        'Industry'
    ]
    for idx, h in enumerate(headers, 2):
        stocks_cells.append(gspread.Cell(row=1, col=idx, value=h))

    row_num = 1
    for stock, data in self.stock_to_data.items():
        row_num += 1
        stocks_cells.append(gspread.Cell(row=row_num,col=1,value='=HYPERLINK("' + str(data['url']) + '","' + str(data['symbol']) + '")'))
        stocks_notes[f'A{row_num}'] = stock
        vals = [
            str(data['LTP']),
            int(data['52W L']),
            int(data['52W H']),
            round((data['LTP']-data['52W L'])/(data['52W H']-data['52W L']) if data.get('52W H') and (data['52W H']-data['52W L'])!=0 else 0,2)*100,
            data['PE'],
            data['Sales_TTM'],
            data['Profit_TTM'],
            data['NPM_TTM'],
            data['Sales_Growth_10Y'],
            data['Sales_Growth_5Y'],
            data['Sales_Growth_3Y'],
            data['Sales_Growth_TTM'],
            data['Profit_Growth_10Y'],
            data['Profit_Growth_5Y'],
            data['Profit_Growth_3Y'],
            data['Profit_Growth_TTM'],
            data['FII_TTM_PTTM'],
            data['Upcoming_Date'] if data['Upcoming_Date'] else data['Reported_Upto'],
            data['Industry']
        ]
        
        for i, val in enumerate(vals, 2):
            stocks_cells.append(gspread.Cell(row=row_num, col=i, value=val))

    stocks.update_cells(stocks_cells, value_input_option='USER_ENTERED')
    stocks.update_notes(stocks_notes)

def run(gc):
  scraper = Screener()
  scraper.get_industry_to_url()
  scraper.get_stock_to_data()
  scraper.update_data_to_sheets(gc)