import pandas as pd
import xlsxwriter
#dataSource
import yfinance as yf

#create a workbook and add a worksheet
workbook = xlsxwriter.Workbook('Stocks.xlsx')
worksheet = workbook.add_worksheet()

#cryptoData
coins = ['BTC-EUR', 'ETH-EUR', 'ADA-EUR', 'VET-EUR']
cryptoData = yf.download(tickers=coins, period='1d', interval='1d')
#clean cryptoData
columns = ["Adj Close", "High", "Low", "Open", "Volume"]
cryptoData.drop(columns=columns, axis=1, inplace=True)
cryptoData = cryptoData.T

#stocksData
stocks = ['AAPL', 'MSFT', 'NEE', 'AMGN', 'TU', 'ZM', 'TSLA', 'TTD', 'SEDG', 'APPN', 'TWLO', 'MMM', 'ABBV', 'SHW', 'K', 'HON', 'AFL', 'JNJ', 'ABT', 'EMR', 'FSLR', 'SPWR', 'ENPH', '^GSPC']
stockNames = ['Apple', 'AbbVie Inc', 'Abbott Labs', 'Aflac', 'Amgen', 'Appian', 'Emerson Electric', 'Enphase  Energy, Inc', 'First Solar, Inc.', 'Honeywell', 'Johnson & Johnson', 'Kellogg Co', '3M Company', 'Microsoft', 'NextEra Energy', 'SolarEdge Technologies', 'Sherwin Williams', 'SunPower Corporation', 'Tesla', 'The Trade Desk', 'Telus', 'Twilio', 'Zoom Video', 'S&P500']
data = yf.download(tickers=stocks, period='1d', interval='1d')
#clean stockData
columns = ["Adj Close", "High", "Low", "Open", "Volume"]
data.drop(columns=columns, axis=1, inplace=True)
data.index = data.index.tz_localize(None)#Remove timezone
data = data.T#Swap Rows/Columns

#appendColumn
data['Stocks'] = stockNames

#concatonateData
df2 = pd.DataFrame(cryptoData)
data = pd.concat([data, df2], ignore_index=False)

#Send data to workbook
datatoexcel = pd.ExcelWriter("Stocks.xlsx", engine='xlsxwriter')
data.to_excel(datatoexcel, sheet_name='Sheet1')
datatoexcel.save()

print(data)
