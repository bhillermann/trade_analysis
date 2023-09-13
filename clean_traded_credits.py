import pandas as pd
from thefuzz import process
import copy
from datetime import datetime

# Import the Traded Credits data with pandas
# Define the Excel file to import
trade_data = ('~/Documents/Trade Analysis/NVCR_Trade-prices_2023-Aug.xlsx')

hu_output_file = ('~/Documents/Trade Analysis/Full-HU-Traded-Credits-{}.csv')
shu_output_file = ('~/Documents/Trade Analysis/Full-SHU-Traded-Credits-{}.csv')

# Open the Excel file. Quit if not FileNotFoundError
try:
    trade_df = pd.ExcelFile(trade_data)
except FileNotFoundError as e:
    print("Excel file not found: ", e)
    exit()

# Grab the HU tab
hu_df = trade_df.parse('Trade prices by HU')

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
    hu_df[x] = hu_df[x].fillna(0)

# Make sure all LTs are integers
hu_df['lt'] = hu_df['lt'].map(int)

choices = ['Corangamite', 'Port Phillip and Westernport', 'Wimmera', \
           'Glenelg Hopkins', 'Goulburn Broken', 'West Gippsland', \
           'East Gippsland', 'Mallee', 'North Central', 'North East']

# Clean up all the inconsistancies in CMA names
def fix_cmas(row):
    cma = process.extractOne(row['cma'], choices)[0]
    return cma

hu_df['cma'] = hu_df.apply(lambda row: fix_cmas(row), axis=1)

hu_df['date'] = pd.to_datetime(hu_df['date'])

# Write the df to a file
hu_df.to_csv(hu_output_file.format(datetime.now().
                                   strftime("%Y%m%d_%H%M%S")))

shu_df.to_csv(shu_output_file.format(datetime.now().\
                                     strftime("%Y%m%d_%H%M%S")))
