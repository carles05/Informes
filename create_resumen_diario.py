import pandas as pd
import numpy as np
import yfinance as yf
import datetime

def tablas_diario(data_i):
    data = data_i.copy()
    weekday_map = {0:'Lun',1:'Mar',2:'Mie',3:'Jue',4:'Vie'}
    data['WeekDay'] = data.WeekDay.map(weekday_map)
    columns = ['Close','Open','Diff%','High','Low','Volume','WeekDay']
    data.index = data.index.strftime('%d/%m/%Y')
    lun = data[data.WeekDay=='Lun'].loc[:,columns]
    mar = data[data.WeekDay=='Mar'].loc[:,columns]
    mie = data[data.WeekDay=='Mie'].loc[:,columns]
    jue = data[data.WeekDay=='Jue'].loc[:,columns]
    vie = data[data.WeekDay=='Vie'].loc[:,columns]
    promedios = pd.DataFrame()
    promedios['%Sube'] = [lun[lun["Diff%"]>0]['Diff%'].count()/lun['Diff%'].count()*100,
                         mar[mar["Diff%"]>0]['Diff%'].count()/mar['Diff%'].count()*100,
                         mie[mie["Diff%"]>0]['Diff%'].count()/mie['Diff%'].count()*100,
                         jue[jue["Diff%"]>0]['Diff%'].count()/jue['Diff%'].count()*100,
                         vie[vie["Diff%"]>0]['Diff%'].count()/vie['Diff%'].count()*100]
    promedios['%Baja'] = [lun[lun["Diff%"]<0]['Diff%'].count()/lun['Diff%'].count()*100,
                         mar[mar["Diff%"]<0]['Diff%'].count()/mar['Diff%'].count()*100,
                         mie[mie["Diff%"]<0]['Diff%'].count()/mie['Diff%'].count()*100,
                         jue[jue["Diff%"]<0]['Diff%'].count()/jue['Diff%'].count()*100,
                         vie[vie["Diff%"]<0]['Diff%'].count()/vie['Diff%'].count()*100]
    promedios['Subida'] = [lun[lun["Diff%"]>0]['Diff%'].mean(),
                                   mar[mar["Diff%"]>0]['Diff%'].mean(),
                                   mie[mie["Diff%"]>0]['Diff%'].mean(),
                                   jue[jue["Diff%"]>0]['Diff%'].mean(),
                                   vie[vie["Diff%"]>0]['Diff%'].mean()]
    promedios['Bajada'] = [lun[lun["Diff%"]<0]['Diff%'].mean(),
                                   mar[mar["Diff%"]<0]['Diff%'].mean(),
                                   mie[mie["Diff%"]<0]['Diff%'].mean(),
                                   jue[jue["Diff%"]<0]['Diff%'].mean(),
                                   vie[vie["Diff%"]<0]['Diff%'].mean()]
    promedios = promedios.T
    return lun, mar, mie, jue, vie, promedios