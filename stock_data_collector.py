import yfinance as yf
from datetime import datetime, date
import gspread

# File path to service_account.json, keep in a folder like "\\%APPDATA%\\gspread\\service_account.json"
file_name = "C:\\Users\\12404\\Documents\\pystuff\\%APPDATA%\\gspread\\service_account.json"

# Name of the Google Sheet file
document_name = "stock_prices"

# The name of the sheet (It's Sheet1 by default)
sheet_name = "Sheet1" 

print("Beginning data collection...")
today = str(date.today())
date_time = str(datetime.now().hour) + ':' + str(datetime.now().minute) + ':' +  str(datetime.now().second)
tsla_price = yf.Ticker('TSLA').info['regularMarketPrice']
aapl_price = yf.Ticker("AAPL").info['regularMarketPrice']
msft_price = yf.Ticker('MSFT').info['regularMarketPrice']
amzn_price = yf.Ticker('AMZN').info['regularMarketPrice']
csco_price = yf.Ticker('CSCO').info['regularMarketPrice']
orcl_price = yf.Ticker('ORCL').info['regularMarketPrice']
goog_price = yf.Ticker('GOOG').info['regularMarketPrice']
nvda_price = yf.Ticker('NVDA').info['regularMarketPrice']
meta_price = yf.Ticker('META').info['regularMarketPrice']
dis_price = yf.Ticker('DIS').info['regularMarketPrice']

stock_dict = [{'ticker': 'TSLA', 'date': today, 'time': date_time, 'price': tsla_price}, 
              {'ticker': 'AAPL', 'date': today, 'time': date_time, 'price': aapl_price},
              {'ticker': 'MSFT', 'date': today, 'time': date_time, 'price': msft_price},
              {'ticker': 'AMZN', 'date': today, 'time': date_time, 'price': amzn_price},
              {'ticker': 'CSCO', 'date': today, 'time': date_time, 'price': csco_price},
              {'ticker': 'ORCL', 'date': today, 'time': date_time, 'price': orcl_price},
              {'ticker': 'GOOG', 'date': today, 'time': date_time, 'price': goog_price},
              {'ticker': 'NVDA', 'date': today, 'time': date_time, 'price': nvda_price},
              {'ticker': 'META', 'date': today, 'time': date_time, 'price': meta_price},
              {'ticker': 'DIS', 'date': today, 'time': date_time, 'price': dis_price}]
header_to_key = {'Ticker': 'ticker', 'Date': 'date', 'Time': 'time', 'Price': 'price'}

print("Data Collected.")
print("Connecting to sheet...")
serv_accnt = gspread.service_account(filename=file_name)
sh = serv_accnt.open(document_name) 
wks = sh.worksheet(sheet_name)

print("Sorting out insertion...")
headers = wks.row_values(1)
put_values = []
for i in stock_dict:
    temp = []
    for j in headers:
        temp.append(i[header_to_key[j]])
    put_values.append(temp)

print("Inserting...")
sh.values_append("Sheet1", {'valueInputOption': 'USER_ENTERED'}, {'values': put_values})
print("Complete.")