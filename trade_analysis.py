#!/usr/bin/env python3

import pandas as pd
from thefuzz import process
import copy
from datetime import datetime, timedelta
import argparse

# Call argparse and define the arguments
parser = argparse.ArgumentParser(description='Process trade prices and supply'
                                 'date to do trade analysis for the past year.')
parser.add_argument("-s", "--supply", required = True, default='supply.xlsx',
                    help='Name of the supply Excel data to read from.')
parser.add_argument("-i", "--input", required = True, 
                    help='The input trade price spreadsheet downloaded from '
                         'the NVCR. "https://www.environment.vic.gov.au/native-'
                         'vegetation/native-vegetation-removal-regulations"')
parser.add_argument("-o", "--output", default='Trade-Analysis.xlsx',
                    help='The name of the file you would like to write the '
                        'anlysis to. Default is "Trade-Analysis.xlsx" in the '
                        'current directory')
parser.add_argument("-b", "--start",
                    help='The date you wish to do the analysis from. '
                     'Default is 12 months ago')
parser.add_argument("-e", "--end",
                    help='The date you wish to do the analysis to. '
                        'Format is YYYY-MM-DD'
                        'Default is the end of the previous month')

args = parser.parse_args()

# Import the Traded Credits data with pandas
# Define the Excel file to import
trade_data = args.input
supply_data = args.supply
output_file = args.output

# Define the property IDs of the Water Authorities
wa = {}
wa['Corangamite'] = ['BBA-2252']
wa['Glenelg Hopkins'] = ['TFN-C0228 ']
wa['Port Phillip and Westernport'] = [
     'BBA-0277', 'BBA-0670', 'BBA-0677', 'BBA-0678']
wa['West Gippsland'] = ['BBA-3049', 'BBA-2845', 'BBA-2839', 'BBA-2790', 
                        'BBA-2789', 'BBA-2751', 'BBA-2766', 'BBA-2623']

# Open the Excel file. Quit if not FileNotFoundError
try:
    trade_df = pd.ExcelFile(trade_data)
except FileNotFoundError as e:
    print("Excel file not found: ", e)
    exit()

try:
    supply_df = pd.read_excel(supply_data, sheet_name=None)
except FileNotFoundError as e:
    print("Excel file not found: ", e)
    exit()

# Grab the HU tab
hu_df = trade_df.parse('Trade prices by HU')

# Rename the columns to something usable
hu_df = hu_df.set_axis([
     'date', 'cma', 'sbv', 'ghu', 'lt', 'sbu', 'ghu_price', 'shu_price', 
     'species', 'price_in_gst', 'price_ex_gst', 'unnamed'], 
     axis=1
     )

# start_date = input('Start date: ')
# end_date = input ('End date: ')

current_date = datetime.today()
end_date = current_date.replace(day=1)
end_date = end_date - timedelta(days=1)
start_date = end_date - timedelta(days=365)

if args.start:
    start_date = datetime.strptime(args.start, '%Y-%m-%d')

if args.end:
    end_date = datetime.strptime(args.end, '%Y-%m-%d')

# start_date = ('2022-06-01')
# end_date = ('2023-05-31')

# start_date = datetime.strptime(start_date, '%Y-%m-%d')
# end_date = datetime.strptime(end_date, '%Y-%m-%d')

# Drop all trades outside of the date range
hu_df = hu_df[((hu_df['date'] >= start_date) & (hu_df['date'] <= end_date))]

# Ensure all 'cma' entries are type string
hu_df['cma'] = hu_df['cma'].map(str)

# Drop the last column because it's not needed
hu_df = hu_df.drop(['unnamed'], axis=1)

# Grab the SHUs from the HU dataframe
shu_df = hu_df[pd.notnull(hu_df['species'])]

# Drop the SHU columns we don't need
shu_df = shu_df.drop(['cma', 'sbv', 'ghu', 'ghu_price'], axis=1)

# Drop the SHU trades so we only have GHU trades
hu_df = hu_df[pd.isnull(hu_df['species'])]

# Drop the HU columns we don't need
hu_df = hu_df.drop(['sbu', 'shu_price', 'species'], axis=1)

# Debug -- write the df to a file
# hu_df.to_excel('~/Documents/output_hu_df.xlsx')

# Make sure all LTs are integers
hu_df['lt'] = hu_df['lt'].map(int)

choices = ['Corangamite', 'Port Phillip and Westernport', 'Wimmera', 
           'Glenelg Hopkins', 'Goulburn Broken', 'West Gippsland', 
           'East Gippsland', 'Mallee', 'North Central', 'North East'
           ]

# Clean up all the inconsistancies in CMA names
def fix_cmas(row):
    cma = process.extractOne(row['cma'], choices)[0]
    return cma

hu_df['cma'] = hu_df.apply(lambda row: fix_cmas(row), axis=1)

# group by CMA and sum GHU, LT and then min, max, median, mean the price
better_df = hu_df.groupby('cma', as_index=False).agg({'ghu': 'sum', 
        'lt': 'sum', 'ghu_price': ['min', 'max', 'median']})

print(better_df, '\n')

summary = {'description': [
                            'Total GHUs traded', 'Total GHUs value',
                            'Average price per GHU',
                            'Median price per GHU', 'Total GHUs without trees',
                            'Total value without trees',
                            'Average price without trees',
                            'Median price without trees', 'Floor price',
                            'Total LTs traded', 'Average LT value',
                            'Supply of Credits', 'Years of Supply', 
                            'LT Supply', 'Water Authority Supply (WA)',
                            'Years of Supply without WA'
                           ], 
                           'values': ['', '', '', '', '', '', '', '', 
                                      '', '', '', '', '', '', '', '']}

summary_df = pd.DataFrame(data=summary)

summaries = dict()

for k, v in hu_df.groupby('cma'):
    # Total GHUs traded
    summary_df.loc[0, ['values']] = v['ghu'].sum()
    # Total GHUs value
    summary_df.loc[1, ['values']] = v['price_ex_gst'].sum()
    # Average price per GHU
    summary_df.loc[2, ['values']] = v['ghu_price'].mean()
    # Median price per GHU
    summary_df.loc[3, ['values']] = v['ghu_price'].median()
    # Total GHUs without trees
    summary_df.loc[4, ['values']] = v.loc[v['lt'] == 0].agg('ghu').sum()
    # Total value without trees
    summary_df.loc[5, ['values']] = v.loc[v['lt'] == 0].agg(
        'price_ex_gst').sum()
    # Average price without trees
    summary_df.loc[6, ['values']] = v.loc[v['lt'] == 0].agg('ghu_price').mean()
    # Median price without trees
    summary_df.loc[7, ['values']] = v.loc[v['lt'] == 0].agg(
         'ghu_price').median()
    # Floor price
    summary_df.loc[8, ['values']] = v['ghu_price'].min()
    # Total LTs traded
    summary_df.loc[9, ['values']] = v['lt'].sum()
    # Calculate the theoretical value of trees
    # Total GHU value - Total GHU without trees value -
    # - Tree GHUs * Avg Price without LTs / Number of LTs
    summary_df.loc[10, ['values']] = (
        (summary_df.loc[1, ['values']]
        - summary_df.loc[5, ['values']]
        - v.loc[v['lt'] > 0].agg('ghu').sum()
        * summary_df.loc[6, ['values']])
        / summary_df.loc[9, ['values']]
        )
    # Supply of Credits
    summary_df.loc[11, ['values']] = supply_df[k].agg('GHU').sum()
    # Years of Supply
    summary_df.loc[12, ['values']] = (summary_df.loc[11, ['values']]
        / summary_df.loc[0, ['values']])
    # LT Supply
    summary_df.loc[13, ['values']] = supply_df[k].agg('LT').sum()
    # Calculate the number of credits owned by water authorities
    wa_credits = 0
    try :
         for x in wa[k]:
            wa_credits = (wa_credits 
                          + supply_df[k].loc[supply_df[k]['Credit Site ID'] 
                               == x].agg('GHU').sum())
    except:
         print(f"No Water Authority credits for {k}.\n")
    # Water Authority Supply (WA)
    summary_df.loc[14, ['values']] = wa_credits
    # Years of Supply without WA
    summary_df.loc[15, ['values']] = ((summary_df.loc[11, ['values']] 
                                      - wa_credits)
                                      / summary_df.loc[0, ['values']])



    summaries[k] = copy.deepcopy(summary_df)


writer = pd.ExcelWriter(output_file.format(datetime.now().\
                                      strftime("%Y%m%d_%H%M%S")), 
                    engine='xlsxwriter',
                    engine_kwargs={'options':{'strings_to_formulas': False}})

hu_df.to_excel(writer, sheet_name='HU Data')
shu_df.to_excel(writer, sheet_name='SHU Data')
hu_df.groupby('cma', as_index=False).agg({'ghu': 'sum', 'lt': 'sum', 
    'ghu_price': ['min', 'max', 'mean', 'median']}).to_excel(
         writer, sheet_name='Summary')

for cma in summaries:
        summaries[cma].to_excel(writer, sheet_name=cma)
        for column in summaries[cma]:
             column_width = max(summaries[cma][column].astype(str).map(len)
                                .max(), len(column))
             col_idx = summaries[cma].columns.get_loc(column)
             if col_idx == 0:
                  column_width = 2
             writer.sheets[cma].set_column(col_idx, col_idx, column_width)

writer.close()
