import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math


#Import S&P500 stock tickers
stocks = pd.read_csv('sp_500_stocks.csv')

#Fetch API data
from secrets import IEX_CLOUD_API_TOKEN

columns = ['Ticker','Price','Mkt Cap', 'Number of Shares to Buy', 'Index %']
final_dataframe = pd.DataFrame(columns=columns)

#Split tickers into size 100 arrays for batch api requests

def divList(lst,n):
    for i in range(0,len(lst),n):
        yield lst[i:i+n]

symbol_groups = list(divList(stocks['Ticker'],100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))


for symbol_string in symbol_strings:

    batch_api_url=f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_url).json()

    for symbol in symbol_string.split(","):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                symbol,
                data[symbol]['quote']['latestPrice'],
                data[symbol]['quote']['marketCap'],
                'N/A',
                'N/A'
            ],
            index = columns
            ),
            ignore_index=True
        )


#Calculate Number of Shares to Buy
portifolio_size = input("Enter the value of your portfolio: ")

try:
    val = float(portifolio_size)
except ValueError:
    print('Please enter a number') 
    portifolio_size = input("Enter the value of your portfolio: ")
    val = float(portifolio_size)

    
total_mkt_cap = final_dataframe['Mkt Cap'].sum()

final_dataframe['Index %'] = final_dataframe['Mkt Cap']/total_mkt_cap

final_dataframe['Number of Shares to Buy'] = final_dataframe['Index %']*val/final_dataframe['Price']

final_dataframe['Number of Shares to Buy'] = final_dataframe['Number of Shares to Buy'].apply(math.floor)

#Save to XLSX
writer = pd.ExcelWriter('recommended trades.xlsx', engine='xlsxwriter')

final_dataframe.to_excel(writer,'Recommended Trades',index=False)

background_color='#0a0a23'
font_color='#ffffff'

string_format = writer.book.add_format(
    {
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format':'$0.00',
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format':'0',
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

float_format = writer.book.add_format(
    {
        'num_format':'0.0000',
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

column_formats={
    'A':['Ticker',string_format],
    'B':['Price',dollar_format],
    'C':['Mkt Cap',integer_format],
    'D':['Number of Shares to Buy',integer_format],
    'E':['Index %',float_format],
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}',18,column_formats[column][1])

writer.save()