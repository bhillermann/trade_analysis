#!/usr/bin/env python3

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
from bs4 import BeautifulSoup
import pandas as pd
import copy
from datetime import datetime

import argparse

# Call argparse and define the arguments
parser = argparse.ArgumentParser(description='Scrape the NVCR for supply data.')
parser.add_argument("-o", "--output", default='Supply_{}.xlsx',
                    help='The name of the file you would like to write the '
                        'supply data to. Default is "Supply.xlsx" in '
                        'the current directory')

args = parser.parse_args()


opts = webdriver.FirefoxOptions()
opts.add_argument("--headless")
driver = webdriver.Firefox(options = opts)

cmas = ['Corangamite', 'Melbourne Water', 'Wimmera', \
           'Glenelg Hopkins', 'Goulburn Broken', 'West Gippsland', \
           'East Gippsland', 'Mallee', 'North Central', 'North East']

all_supply = dict()
big_df = pd.DataFrame()

supply_xlsx = args.output
supply_csv = '~/Documents/Trade Analysis/Supply_{}.csv'

for x in cmas:
    print('Scraping supply data for:', x, '...\n')
    driver.get("https://nvcr.delwp.vic.gov.au/Search/GHU")
    time.sleep(2)
    
    ghu_element = driver.find_element(By.XPATH, 
            '//*[@id="GeneralGuidelineSearch"]/div[2]/div[1]/div/input')
    sbv_element = driver.find_element(By.XPATH, 
            '//*[@id="GeneralGuidelineSearch"]/div[2]/div[2]/div/input')
    lt_element = driver.find_element(By.XPATH, 
            '//*[@id="GeneralGuidelineSearch"]/div[2]/div[3]/div/input')
    search_button = driver.find_element(By.XPATH, 
            '//*[@id="GeneralGuidelineSearch"]/div[2]/div[7]/div[2]/button')

    cma_select = Select(driver.find_element(By.XPATH, 
            r'//*[@id="GeneralGuidelineSearch"]/div[2]/div[5]'
            r'/div/table/tbody/tr/td[2]/div[2]/select'))

    ghu_element.send_keys("0.001")
    sbv_element.send_keys("0.001")
    lt_element.send_keys("0")

    cma_select.select_by_value(x)

    search_button.click()

    time.sleep(5)

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find_all("table", {"class":"table"})
    all_tables = pd.read_html(str(div))
    supply_table = all_tables[4]
    all_supply[x] = copy.deepcopy(all_tables[4])
    big_df = pd.concat([big_df, supply_table], ignore_index=True)

driver.quit()

big_df['Date'] = datetime.now().strftime("%Y-%m-%d")

with pd.ExcelWriter(supply_xlsx.format(datetime.now().\
                                      strftime("%Y%m%d_%H%M%S")), 
                    engine='xlsxwriter', \
                    engine_kwargs={'options':{'strings_to_formulas': False}})\
                    as writer:
    for cma in all_supply:
        all_supply[cma].to_excel(writer, sheet_name=cma)

big_df.to_csv(supply_csv.format(datetime.now().\
                                      strftime("%Y%m%d_%H%M%S")))