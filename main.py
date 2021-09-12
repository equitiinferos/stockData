import pandas as pd
import xlsxwriter
from binance.client import Client

#Create a workbook and add a worksheet
workbook = xlsxwriter.Workbook('Stocks.xlsx')
worksheet = workbook.add_worksheet()

#Data Source
import yfinance as yf

#stocksData
stocks = ['AAPL', 'MSFT', 'NEE', 'AMGN', 'TU', 'ZM', 'TSLA', 'TTD', 'SEDG', 'APPN', 'TWLO', 'MMM', 'ABBV', 'SHW', 'K', 'HON', 'AFL', 'JNJ', 'ABT', 'EMR', 'FSLR', 'SPWR', 'ENPH']
stockNames = ['Apple', 'Microsoft', 'NextEra Energy', 'Amgen', 'Telus', 'Zoom Video', 'Tesla', 'The Trade Desk', 'SolarEdge Technologies', 'Appian', 'Twilio', '3M Company', 'AbbVie Inc', 'Sherwin Williams', 'Kellogg Co', 'Honeywell', 'Aflac', 'Johnson & Johnson', 'Abbott Labs', 'Emerson Electric', 'First Solar, Inc.', 'SunPower Corporation',  'Enphase  Energy, Inc' ]
data = yf.download(tickers=stocks, period='1d', interval='1d')

print(data)

#clean data
columns = ["Adj Close", "High", "Low", "Open", "Volume"]
data.drop(columns=columns, axis=1, inplace=True)

#remove timezone
data.index = data.index.tz_localize(None)

#swap Rows/Columns
data = data.T

#append column
data[""] = ""
data['Stocks'] = stockNames

#send stocksData to excel
datatoexcel = pd.ExcelWriter("Stocks.xlsx", engine='xlsxwriter')
data.to_excel(datatoexcel, sheet_name='Sheet1')
datatoexcel.save()
