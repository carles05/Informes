import pandas as pd
import numpy as np
import yfinance as yf
import datetime

def notas_tables(data_i):
    data = data_i.copy()
    data['FDM'] = data.Month!=data.Month.shift(1) #Last Day of Month
    data['PDM'] = data.Month!=data.Month.shift(-1) #First Day of Month
    data['Close_ant'] = data.Close.shift(-1)
    years = data.Year.unique()[0:3]
    notas_tables={}
    for year in years:
        data_t = data[data.Year==year]
        maxims = data_t.groupby(by='Month')['High'].max()
        minims = data_t.groupby(by='Month')['Low'].min()
        inicis = data_t[data_t.PDM==True].groupby(by='Month')['Close_ant'].sum()
        perc_max = 100*(maxims-inicis)/inicis
        perc_min = 100*(minims-inicis)/inicis
        df = pd.DataFrame([inicis,maxims,perc_max,minims,perc_min]).T
        df.columns = ['Inici','Maxim','%Max','Minim','%Min']
        #df.index = ['Gener','Febrer','Març','Abril','Maig','Juny','Juliol','Agost','Septembre','Octubre','Novembre','Decembre']
        months_map = {1:'Gener',2:'Febrer',3:'Març',4:'Abril',5:'Maig',6:'Juny',7:'Juliol',
         8:'Agost',9:'Septembre',10:'Octubre',11:'Novembre',12:'Decembre'}
        df.index = data_t.Month.unique()[::-1]
        df.index = df.index.map(months_map)
        notas_tables[year] = df
    return notas_tables
