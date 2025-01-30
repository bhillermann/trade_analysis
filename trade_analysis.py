#!/usr/bin/env python3

import pandas as pd
import numpy as np
from thefuzz import process
import copy
from datetime import datetime, timedelta
import argparse
from ghu_search import get_supply
from download_nvcr import get_trade_data
from openpyxl import load_workbook
from openpyxl.styles import Font

# Call argparse and define the arguments
parser = argparse.ArgumentParser(description='Process trade prices and supply'
                                 'date to do trade analysis for the past '
                                 'year.')
parser.add_argument("-s", "--supply", required = False,
                    help='Name of the supply Excel data to read from. Default'
                    ' is to scrape new data.')
parser.add_argument("-i", "--input", required = False, 
                    help='The input trade price spreadsheet downloaded from '
                         'the NVCR. "https://www.environment.vic.gov.au/'
                         'native-vegetation/native-vegetation-removal-'
                         'regulations". '
                         'Not using this switch will download a new file.')
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
output_file = args.output
download_trade_data = False

if args.supply == None:
    print('Downloading supply data...\n\n')
    supply_data = 'Supply_{}.xlsx'.format(datetime.now().
                                          strftime("%Y%m%d_%H%M%S"))
    get_supply(supply_data, False)
    print('Supply data saved as: ', supply_data)
else:
    supply_data = args.supply

if args.input == None:
    print('Downloading NVCR trade data...')
    trade_data = 'NVCR_Trade-prices-{}.xlsx'.format(datetime.now().
                                                    strftime("%Y%m%d_%H%M%S"))
    get_trade_data(trade_data)
    print('Trade data saved as: ', trade_data)
else:
    trade_data = args.input


# Define the property IDs of the Water Authorities
wa = {}
wa['Corangamite'] = ['BBA-2252']
wa['Glenelg Hopkins'] = ['TFN-C0228 ']
wa['Melbourne Water'] = [
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
hu_df = trade_df.parse('Trade Prices by HU')

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

# Ensure all 'cma' entries are type string
hu_df['cma'] = hu_df['cma'].map(str)

# Change Date from datetime to date
hu_df['date']=hu_df['date'].dt.date

# Drop the last column because it's not needed
hu_df = hu_df.drop(['unnamed'], axis=1)

# Grab the SHUs from the HU dataframe
shu_df = hu_df[pd.notnull(hu_df['species'])]

# Drop all HU trades outside of the date range
hu_df = hu_df[((hu_df['date'] >= start_date.date()) & (hu_df['date'] <= end_date.date()))]

# Drop all SHU trades outside of a three year period from end date
shu_df = shu_df[((shu_df['date'] >= end_date.date() - timedelta(days=1095)) & 
                 (shu_df['date'] <= end_date.date()))]

# Drop the SHU columns we don't need
shu_df = shu_df.drop(['cma', 'sbv', 'ghu', 'ghu_price'], axis=1)

# Create Summary SHU Data
shu_summary = {'Description': [
                                'Number of SHU trades', 
                                'Total SHUs traded' ,
                                'Total Value of SHU trades', 
                                'Average Price per SHU',
                                'SHU Floor Price',
                                'SHU Ceiling Price',
                                'SHU median price'
                            ], 
                           'Values': [
                                (shu_df.groupby(['date', 'shu_price'])
                                 .sum('sbu')['sbu'].count()),
                                shu_df['sbu'].sum(), 
                                shu_df['price_ex_gst'].sum(), 
                                (shu_df['price_ex_gst'].sum()
                                 / shu_df['sbu'].sum()), 
                                shu_df['shu_price'].min(), 
                                shu_df['shu_price'].max(), 
                                np.median(shu_df['shu_price'].unique())
                            ]}

shu_summary_df = pd.DataFrame(shu_summary)

# Drop the SHU trades so we only have GHU trades
hu_df = hu_df[pd.isnull(hu_df['species'])]

# Drop the HU columns we don't need
hu_df = hu_df.drop(['sbu', 'shu_price', 'species'], axis=1)

# Make sure all LTs are integers
hu_df['lt'] = hu_df['lt'].map(int)

cmas = ['Corangamite', 'Melbourne Water', 'Port Phillip and Westernport', 
           'Wimmera', 'Glenelg Hopkins', 'Goulburn Broken', 'West Gippsland', 
           'East Gippsland', 'Mallee', 'North Central', 'North East'
           ]

# Clean up all the inconsistancies in CMA names
def fix_cmas(row):
    cma = process.extractOne(row['cma'], cmas)[0]
    return cma

hu_df['cma'] = hu_df.apply(lambda row: fix_cmas(row), axis=1)
hu_df = hu_df.replace('Port Phillip and Westernport', 'Melbourne Water')

# This needs to be cleaned up. Use normal headers and then relable after 
# calculations
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

print('Calculating per CMA data-------------------------------------------\n')
for k, v in hu_df.groupby('cma'):
    print(f'Crunching data for {k}...\n')
    # Total GHUs traded
    summary_df.loc[0, ['values']] = v['ghu'].sum()
    # Total GHUs value
    summary_df.loc[1, ['values']] = v['price_ex_gst'].sum()
    # Average price per GHU
    summary_df.loc[2, ['values']] = v['price_ex_gst'].sum() / v['ghu'].sum()
    # Median price per GHU
    summary_df.loc[3, ['values']] = v['ghu_price'].median()
    # Total GHUs without trees
    summary_df.loc[4, ['values']] = v.loc[v['lt'] == 0].agg('ghu').sum()
    # Total value without trees
    summary_df.loc[5, ['values']] = v.loc[v['lt'] == 0].agg(
        'price_ex_gst').sum()
    # Average price without trees
    summary_df.loc[6, ['values']] = (
            v.loc[v['lt'] == 0].agg('price_ex_gst').sum() 
            / v.loc[v['lt'] == 0].agg('ghu').sum()
        )
    # Median price without trees
    summary_df.loc[7, ['values']] = v.loc[v['lt'] == 0].agg(
         'ghu_price').median()
    # Floor price
    summary_df.loc[8, ['values']] = v['ghu_price'].min()
    # Total LTs traded
    summary_df.loc[9, ['values']] = v['lt'].sum()
    # Calculate the theoretical value of trees
    # (Total GHU value - ((Total GHUs - Total GHUs without trees)
    # * Avg price without trees) - Total value without trees)
    # / Total LTs Traded
    summary_df.loc[10, ['values']] = (
        (
            summary_df.loc[1, ['values']]
            - (
                (
                    summary_df.loc[0, ['values']]
                    - summary_df.loc[4, ['values']]
                )
                * summary_df.loc[6, ['values']]
            )  
            - summary_df.loc[5, ['values']]
        )
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


# Writing it all to Excel

print('Creating Excel Spreadsheet...\n\n')

writer = pd.ExcelWriter(output_file, 
                    engine='xlsxwriter',
                    engine_kwargs={'options':{'strings_to_formulas': False}})

workbook = writer.book

# Define the different formats

currency_format = workbook.add_format(
    {
        'num_format': '$#,##0.00'
    }
)

# Writing the HU information ------------------------------------------------
sheetname = 'HU Data'
# Write the HU dataframe to sheet HU Data
hu_df.to_excel(writer, sheet_name=sheetname, 
               startrow=1, header=False, index=False)

# Create some human readable headers
header = ('Date', 'CMA', 'SBV', 'GHU', 'LT', 'GHU Price', 
             'Price (in GST)', 'Price (ex GST)') 

column_settings = [{"header": column} for column in header]

# Get the dimensions of the dataframe.
(max_row, max_col) = hu_df.shape

# Set the active sheet to HU Data
worksheet = writer.sheets[sheetname]

# Add the Excel table structure. Pandas added the data.
worksheet.add_table(0, 0, max_row, max_col - 1, 
                    {
                        'columns': column_settings,
                        'style': 'Table Style Light 11',
                        'banded_columns': True
                    })

worksheet.set_column(max_col-3, max_col - 1, None, currency_format)

worksheet.autofit()

# End HU dataframe ----------------------------------------------------------

# Start - Writing SHU information to file -----------------------------------
sheetname = 'SHU Data'
# Write the SHU dataframe to sheet SHU Data
shu_df.to_excel(writer, sheet_name=sheetname, startrow=1, 
                header=False, index=False)

# Create some human readable headers
header = ('Date', 'LT',	'SHUs',	'SHU Price', 'Species', 
              'Price (in GST)', 'Price (ex GST)') 

column_settings = [{"header": column} for column in header]

# Get the dimensions of the dataframe.
(max_row, max_col) = shu_df.shape

# Set the active sheet to SHU Data
worksheet = writer.sheets[sheetname]

# Add the Excel table structure. Pandas added the data.
worksheet.add_table(0, 0, max_row, max_col - 1, 
                    {
                        'columns': column_settings,
                        'style': 'Table Style Light 11',
                        'banded_columns': True
                    })

# Set currency format on pricing columns
worksheet.set_column(max_col-2, max_col - 1, None, currency_format)
worksheet.set_column(3, 3, None, currency_format)

# Write the SHU Summary data
shu_summary_df.to_excel(writer, sheet_name=sheetname, 
                        startrow=1, startcol=8, index=False, header=False)

# Get the dimensions of the dataframe.
(max_row, max_col) = shu_summary_df.shape


# Add the Excel table structure. Pandas added the data.
worksheet.add_table(1, 8, max_row, max_col +8 - 1, 
                    {
                        'style': 'Table Style Light 18',
                        'autofilter': False,
                        'header_row': False,
                        'first_column': True
                    })

# Autofit columns
worksheet.autofit()
# End SHU dataframe ---------------------------------------------------------

# Overview Summary Dataframe ------------------------------------------------
sheetname = 'HU Summary'
# Create high level summary data
hu_summary = hu_df.groupby('cma', as_index=False).agg(
    {
        'ghu': 'sum', 'lt': 'sum', 'price_ex_gst': 'sum', 
        'ghu_price': ['min', 'max', 'mean', 'median']
    }
    )

# Write it to Excel 
hu_summary.to_excel(writer, sheet_name=sheetname, header=False)

# Create some human readable headers
header = ('Index', 'CMA', 'GHUs', 'LTs', 'Total Value', 'GHU Floor Price', 
                     'GHU Ceiling Price', 'GHU Mean', 'GHU Median', 
                     'GHU Weighted Average') 

column_settings = [{"header": column} for column in header]

# Get the dimensions of the dataframe.
(max_row, max_col) = hu_summary.shape

# Set the active sheet to SHU Data
worksheet = writer.sheets[sheetname]

# Add the Excel table structure. Pandas added the data.
worksheet.add_table(0, 0, max_row, max_col + 1, 
                    {
                        'columns': column_settings,
                        'style': 'Table Style Light 11',
                        'banded_columns': True
                    })

# Set currency format on pricing columns
worksheet.set_column(max_col - 4, max_col + 1, None, currency_format)

# Autofit columns
worksheet.autofit()

# End Overview Summary data -------------------------------------------------

for cma in summaries:
        summaries[cma].columns = ['Description', 'Values']
        summaries[cma].to_excel(writer, sheet_name=cma, index=False)
        writer.sheets[cma].autofit()

        # Get the dimensions of the dataframe.
        (max_row, max_col) = summaries[cma].shape

        worksheet = writer.sheets[cma]

        # Create some human readable headers
        header = ('Metric', 'Value') 

        column_settings = [{"header": column} for column in header]

        # Add the Excel table structure. Pandas added the data.
        worksheet.add_table(0, 0, max_row, max_col - 1, 
                            {
                                'columns': column_settings,
                                'style': 'Table Style Light 11',
                                'banded_columns': True,
                                'autofilter': False
                            })

writer.close()

# Complete the formatting that can't be done by XlsxWriter. Use openpyxl

print('Putting final touches on formatting...\n\n')

workbook = load_workbook(filename=output_file)

# Set the font format we want to use
default_font = Font(name='Rubik Light', size=10)

# Set the currenct formate
currency_format = '$#,##0.00'

# Iterate over all cells in all sheets and set the font
for x in workbook.sheetnames:
    sheet = workbook[x]
    for row in sheet["A:J"]:
        for cell in row:
            cell.font = default_font

# Set the weighted average formula on the CMA Summary page
sheet = workbook['HU Summary']

# Define which cells need the formula
for row in sheet["J2:J11"]:
    for cell in row:
        cell.value = "=E{row}/C{row}".format(row=cell.row)
        cell.number_format = currency_format

# Set the currency format on the summary table in SHU Data tab --------------
sheet = workbook['SHU Data']

# Define which cells in the summary table need to be set to currency
for row in sheet["J4:J8"]:
    for cell in row:
        cell.number_format = currency_format


# Update the cmas to reflect Melbourne Water before we iterate through 
# summary pages

print(f'{cmas}\n')
cmas.remove('Port Phillip and Westernport')
print(f'{cmas}\n')

# Define which cells on the CMA pages need to be set to currency
currency_cells = ('B3', 'B4', 'B5', 'B7', 'B8', 'B9', 'B10', 'B12')

# Iterate over the CMA sheets and set the currency format
for x in cmas:
    sheet = workbook[x]
    for cells in currency_cells:
        sheet[cells].number_format = currency_format
    
workbook.save(filename=output_file)

print('Analyses complete.')
