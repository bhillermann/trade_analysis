#!/usr/bin/env python3

from typing import Any

import pandas as pd
from thefuzz import process
from datetime import datetime
import argparse

# Call argparse and define the arguments
parser = argparse.ArgumentParser(description='Process NVCR trading information'
                                 'to a clean CSV for importing into other'
                                 'systems.')

parser.add_argument("-i", "--input", required = True, 
                    help='The input trade price spreadsheet downloaded from '
                         'the NVCR. "https://www.environment.vic.gov.au/native-'
                         'vegetation/native-vegetation-removal-regulations"')

args = parser.parse_args()

# Import the Traded Credits data with pandas
# Define the Excel file to import
trade_data = args.input

hu_output_file = ('~/Documents/Trade Analysis/Full-HU-Traded-Credits-{}.csv')
shu_output_file = ('~/Documents/Trade Analysis/Full-SHU-Traded-Credits-{}.csv')

# Open the Excel file. Quit if not FileNotFoundError
try:
    trade_df = pd.ExcelFile(trade_data)
except FileNotFoundError as e:
    print("Excel file not found: ", e)
    exit()

# Grab the HU tab
hu_df = trade_df.parse('Trade Prices by HU')

# Rename the columns to something usable
hu_df = hu_df.set_axis(['date', 'cma', 'sbv', 'ghu', 'lt', 'sbu', 'ghu_price',\
                'shu_price', 'species', 'price_in_gst', 'price_ex_gst',\
                'unnamed'], axis=1)

# Ensure all 'cma' entries are type string
hu_df['cma'] = hu_df['cma'].map(str)

# Grab the SHUs from the HU dataframe
shu_df = hu_df[pd.notnull(hu_df['species'])]

# Drop the SHU columns we don't need
shu_df = shu_df.drop(['cma', 'sbv', 'ghu', 'ghu_price', 'unnamed'], axis=1)

# Drop the SHU trades so we only have GHU trades
hu_df = hu_df[pd.isnull(hu_df['species'])]

# Drop the columns we don't need
hu_df = hu_df.drop(['unnamed', 'sbu', 'shu_price', 'species'], axis=1)


# Replace any NaN values with 0
for x in ['sbv', 'ghu', 'lt', 'ghu_price', 'price_in_gst', 'price_ex_gst']:
    hu_df[x] = hu_df[x].infer_objects(copy=False).fillna(0) # Have to infer
    # objects as downcasting behaviour is deprecated 

# Make sure all LTs are integers
hu_df['lt'] = hu_df['lt'].map(int)

choices = ['Corangamite', 'Port Phillip and Westernport', 'Melbourne Water',
           'Wimmera', 'Glenelg Hopkins', 'Goulburn Broken', 'West Gippsland',
           'East Gippsland', 'Mallee', 'North Central', 'North East']

# Clean up all the inconsistancies in CMA names
def fix_cmas(row: pd.Series[Any]) -> str:
    result = process.extractOne(row['cma'], choices)  # type: ignore[attr-defined]
    if result is None:
        return str(row['cma'])
    return str(result[0])

hu_df['cma'] = hu_df.apply(lambda row: fix_cmas(row), axis=1)

# Change PPWP to Melbourne Water
hu_df = hu_df.replace('Port Phillip and Westernport', 'Melbourne Water')

hu_df['date'] = pd.to_datetime(hu_df['date'])

# Write the df to a file
hu_df.to_csv(hu_output_file.format(datetime.now().
                                   strftime("%Y%m%d_%H%M%S")))

shu_df.to_csv(shu_output_file.format(datetime.now().\
                                     strftime("%Y%m%d_%H%M%S")))
