import requests
from bs4 import BeautifulSoup as bs
from datetime import datetime
import pandas as pd

nvcr = ('https://www.environment.vic.gov.au/'
        'native-vegetation/native-vegetation-removal-regulations')

resp = requests.get(nvcr)
soup = bs(resp.text, 'lxml')
download_link = ""

for link in soup.find_all('a'):
    x = str(link).encode('utf8').find(b'Traded credits information')
    if x != -1:
        print("I'm here!!\n\n----------")
        download_link = link.get('href')

print('This is the download link: ', download_link, '\n\n')

response = requests.get(download_link)
print(response.status_code)

filename = ('/home/bhillermann/Documents/Trade Analysis/'
            'NVCR_Trade-prices-{}.xlsx'.format(datetime.now()
                                     .strftime("%Y%m%d_%H%M%S")))

with open(filename, 'wb') as file:
    file.write(response.content)

trade_df = pd.ExcelFile(filename)
# Grab the HU tab
hu_df = trade_df.parse('Trade Prices by HU')
print(hu_df)