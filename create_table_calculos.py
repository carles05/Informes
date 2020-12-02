import pandas as pd
import numpy as np
import yfinance as yf
import datetime

def calculos_table(data, semanas=20):
    data_f = data.copy()
    data_f['LDW'] = data_f.Week != data_f.Week.shift(1)
    data_f["Date"] = data_f.index
    maxset = data_f.groupby(by = ["Year","Week"], as_index=False)["High"].max()
    minset = data_f.groupby(by = ["Year","Week"])["Low"].min()
    data_f = data_f[data_f["LDW"]==True]
    data_f = data_f.head(semanas+15)
    data_f = pd.merge(data_f,maxset,how='left',on=['Year','Week'])
    data_f = pd.merge(data_f,minset,how='left',on=['Year','Week'])
    data_f.rename(columns = {'High_x':'High', 'High_y':'Max. Set',
                             'Low_x':'Low', 'Low_y':'Min. Set'}, inplace=True)
    data_f["Dif. Max i Min"] = data_f['Max. Set']-data_f['Min. Set']
    data_f.index = data_f.Date
    data_f.drop(columns='Date', inplace=True)
    data_f["Variacio Setmanal"] = 100*(data_f.Close/data_f.Close.shift(-1)-1)
    data_f["VariacioAbs"] = data_f['Variacio Setmanal'].apply(abs)
    data_f["BB Superior"] = pd.DataFrame([data_f.Close,data_f.Close.shift(-1),data_f.Close.shift(-2),data_f.Close.shift(-3),
                       data_f.Close.shift(-4),data_f.Close.shift(-5),data_f.Close.shift(-6),data_f.Close.shift(-7),
                       data_f.Close.shift(-8),data_f.Close.shift(-9)]).std()+data_f.Close
    data_f["BB Inferior"] = data_f.Close-pd.DataFrame([data_f.Close,data_f.Close.shift(-1),data_f.Close.shift(-2),
                       data_f.Close.shift(-3),data_f.Close.shift(-4),data_f.Close.shift(-5),data_f.Close.shift(-6),
                       data_f.Close.shift(-7),data_f.Close.shift(-8),data_f.Close.shift(-9)]).std()
    data_f["Porc. Var. 10"] = pd.DataFrame([data_f.VariacioAbs,data_f.VariacioAbs.shift(-1),data_f.VariacioAbs.shift(-2),
                                      data_f.VariacioAbs.shift(-3),data_f.VariacioAbs.shift(-4),
                                      data_f.VariacioAbs.shift(-5),data_f.VariacioAbs.shift(-6),
                                      data_f.VariacioAbs.shift(-7),data_f.VariacioAbs.shift(-8),
                                      data_f.VariacioAbs.shift(-9)]).mean()
    data_f["2a Pujada"] = (data_f['Porc. Var. 10']/100*2)*data_f.Close+data_f.Close
    data_f["1a Pujada"] = (data_f['Porc. Var. 10']/100)*data_f.Close+data_f.Close
    data_f["2a Baixada"] = data_f.Close-(data_f['Porc. Var. 10']/100*2)*data_f.Close
    data_f["1a Baixada"] = data_f.Close-(data_f['Porc. Var. 10']/100)*data_f.Close
    data_f["Dif preu ant a Max"] = data_f['Max. Set']-data_f['Close'].shift(-1)
    data_f["Dif preu ant a Min"] = data_f['Min. Set']-data_f['Close'].shift(-1)
    columns = ['Variacio Setmanal', 'VariacioAbs','BB Superior','BB Inferior','2a Pujada','1a Pujada',
              '2a Baixada','1a Baixada','Porc. Var. 10','Close','Max. Set','Min. Set','Dif. Max i Min',
              "Dif preu ant a Max","Dif preu ant a Min"]
    data_f = data_f[columns]
    var_set = data_f['Variacio Setmanal'].head(semanas)
    promedios_calc = {'%Sube': var_set[var_set>0].count()/var_set.count()*100,
                     '%Baja': var_set[var_set<0].count()/var_set.count()*100,
                     'Promedio Subida': var_set[var_set>0].mean(),
                     'Promedio Bajada': var_set[var_set<0].mean()}
    ind = str(semanas)+' Setmanes'
    promedios_calc = pd.DataFrame(promedios_calc, index=[ind]).T
    return data_f.head(semanas), promedios_calc
