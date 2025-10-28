#!/usr/bin/env python3

import requests
from bs4 import BeautifulSoup as bs
from datetime import datetime

nvcr = ('https://www.environment.vic.gov.au/'
        'native-vegetation/native-vegetation-removal-regulations')

def get_trade_data(filename):
    print("Running get_trade_data() now...")
    resp = requests.get(nvcr)
    soup = bs(resp.text, 'lxml')
    download_link = ""

    for link in soup.find_all('a'):
        x = str(link).encode('utf8').find(b'Traded credits information')
        if x != -1:
            download_link = link.get('href')

    print('This is the download link: ', download_link, '\n\n')

    response = requests.get(download_link)
    print(response.status_code)

    with open(filename, 'wb') as file:
        file.write(response.content)

if __name__ == "__main__":

    filename = ('/home/bhillermann/Documents/Trade Analysis/'
                'NVCR_Trade-prices-{}.xlsx'.format(datetime.now()
                                        .strftime("%Y%m%d_%H%M%S")))
    
    get_trade_data(filename)