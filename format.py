from openpyxl import load_workbook
from openpyxl.styles import DEFAULT_FONT
from openpyxl.styles import Font

workbook = load_workbook(filename='/home/bhillermann/Documents/Trade Analysis/Latest_trade.xlsx')
sheet = workbook.active

default_font = Font(name='Rubik Light', size=10)


for x in workbook.sheetnames:
    sheet = workbook.get_sheet_by_name(x)
    for row in sheet["A:J"]:
        for cell in row:
            cell.font = default_font    

cmas = ['Corangamite', 'Melbourne Water', 'Wimmera', 'Glenelg Hopkins', 
           'Goulburn Broken', 'West Gippsland', 'East Gippsland', 'Mallee', 
           'North Central', 'North East'
           ]

currency_format = '$#,##0.00'
currency_cells = ('B3', 'B4', 'B5', 'B7', 'B8', 'B9', 'B10', 'B12')

for x in cmas:
    sheet = workbook.get_sheet_by_name(x)
    print(sheet.title)
    for cells in currency_cells:
        sheet[cells].number_format = currency_format
    



# print(sheet.title)


DEFAULT_FONT.name = 'Arial'
DEFAULT_FONT.size = 10

{k: setattr(DEFAULT_FONT, k, v) for k, v in default_font.__dict__.items()}

workbook.save(filename='/home/bhillermann/Documents/Trade Analysis/Latest_trade_updated.xlsx')