#!/usr/bin/env python3

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
from bs4 import BeautifulSoup
import pandas as pd
import copy
from datetime import datetime
import argparse
from io import StringIO


def get_supply() -> dict[str, pd.DataFrame]:
    opts = webdriver.FirefoxOptions()
#    opts.add_argument("--headless")
    driver = webdriver.Firefox(options = opts)
    wait = WebDriverWait(driver, timeout=10)

    cmas = ['Corangamite', 'Melbourne Water', 'Wimmera', \
            'Glenelg Hopkins', 'Goulburn Broken', 'West Gippsland', \
            'East Gippsland', 'Mallee', 'North Central', 'North East']

    all_supply: dict[str, pd.DataFrame] = {}

    for x in cmas:
        print('Scraping supply data for:', x, '...\n')
        driver.get("https://nvcr.delwp.vic.gov.au/Search/GHU")
        
        wait.until(EC.element_to_be_clickable((By.XPATH, 
                '//*[@id="GeneralGuidelineSearch"]/div[2]/div[1]/div/input')))
        
        ghu_element = driver.find_element(By.XPATH, 
                '//*[@id="GeneralGuidelineSearch"]/div[2]/div[1]/div/input')
        sbv_element = driver.find_element(By.XPATH, 
                '//*[@id="GeneralGuidelineSearch"]/div[2]/div[2]/div/input')
        lt_element = driver.find_element(By.XPATH, 
                '//*[@id="GeneralGuidelineSearch"]/div[2]/div[3]/div/input')
        search_button = driver.find_element(By.XPATH, 
                '//*[@id="GeneralGuidelineSearch"]/div[2]/div[7]/div[2]/button')

        cma_select = Select(driver.find_element(By.XPATH, 
                '//*[@id="GeneralGuidelineSearch"]/div[2]/div[5]'
                '/div/table/tbody/tr/td[2]/div[2]/select'))

        ghu_element.send_keys("0.001")
        sbv_element.send_keys("0.001")
        lt_element.send_keys("0")

        cma_select.select_by_value(x)

        search_button.click()

        wait.until(EC.element_to_be_clickable((By.XPATH,
                '/html/body/div[3]/div[1]/div[3]/div[7]/div[3]/label')))

        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        div = soup.find_all("table", {"class":"table"})
        all_tables = pd.read_html(StringIO(str(div)))
        all_supply[x] = copy.deepcopy(all_tables[4])

    driver.quit()

    return all_supply


if __name__ == "__main__":

    # Call argparse and define the arguments
    parser = argparse.ArgumentParser(description='Scrape the NVCR for supply data.')
    parser.add_argument("-o", "--output", default='Supply_{}.xlsx',
                        help='The name of the file you would like to write the '
                            'supply data to. Default is "Supply_{timestamp}.xlsx" in '
                            'the current directory')

    args = parser.parse_args()

    # Get supply data as dict of DataFrames
    all_supply = get_supply()

    # Write to Excel file
    supply_xlsx = args.output.format(datetime.now().strftime("%Y%m%d_%H%M%S"))
    with pd.ExcelWriter(supply_xlsx,
                        engine='xlsxwriter',
                        engine_kwargs={'options':{'strings_to_formulas': False}}) as writer:
        for cma in all_supply:
            all_supply[cma].to_excel(writer, sheet_name=cma)

    print(f'Supply data saved to: {supply_xlsx}')


