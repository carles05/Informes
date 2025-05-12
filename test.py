import yfinance as yf
import pandas as pd

isins = ['IE0004445015',
         'LU1006075227',
         'LU1442550031',
         'LU1006080144',
         'LU0784385337',
         'LU0957039414',
         'LU1442550031',
         'LU0638558980']

index_ticker = yf.Ticker(isins[2])
index_data = pd.DataFrame(index_ticker.history(period="max"))