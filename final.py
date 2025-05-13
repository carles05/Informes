# One file Code
import numpy as np
import yfinance as yf
import pandas as pd
from datetime import date
import math
import time
import warnings
warnings.filterwarnings("ignore")

try:
    # Try to import curl_cffi to handle Yahoo Finance rate limits
    from curl_cffi import requests
    USE_CURL_CFFI = True
except ImportError:
    USE_CURL_CFFI = False
    print("curl_cffi not installed. Using standard requests instead.")
    print("To avoid rate limits, install curl_cffi: pip install curl_cffi")

def controlIndex(ind):
    today = date.today()
    stock_data = read_data(ind)
    index_metrics = pd.DataFrame()
    # Posicio Actual
    index_metrics['Posicio Actual'] = [stock_data.Close[0]]

    # Variacio Setmanal 10 Setmanes
    tabla_calculos, promedios_calc = calculos_table(stock_data, semanas=60)
    index_metrics['Variacio 10 Setmanes'] = [round(tabla_calculos['Porc. Var. 10'][0],2)]
    
    # Percentatge Banda Boellinger
    index_metrics['Perc. BBS'] = [round((tabla_calculos['BB Superior'][0]/tabla_calculos['Close'][0]-1)*100,2)]

    # Revalorització setmana actual
    index_metrics['Revaloritzacio Setmana Actual'] = [round(tabla_calculos['Variacio Setmanal'][0],2)]

    # Revaloritzacio mes actual
    months = months_data(stock_data)
    index_metrics['Revaloritzacio Mes Actual'] = [round(months['Diff%'][0],2)]

    # Mitja 52 setmanes: (tancament/mitja52-1)*100
    mitja52 = tabla_calculos['Close'][::-1].rolling(window=52).mean()[-1]
    index_metrics['Mitja 52 Setmanes respecte actual'] = [round((stock_data.Close[0]/mitja52-1)*100,2)]
    
    # Cotitzacions diaries consecutives
    if stock_data['Diff%'][0]>0:
        c = 0
        while sum(stock_data['Diff%'].head(c+1)>0)==c+1:
            c+=1
        index_metrics['Cotitzacions Consecutives diaries'] = [c]
    else:
        c = 0
        while sum(stock_data['Diff%'].head(c+1)<0)==c+1:
            c+=1
        index_metrics['Cotitzacions Consecutives diaries'] = [-c]
    
    # Cotitzacions setmanals consecutives
    if tabla_calculos['Variacio Setmanal'][0]>0:
        c = 0
        while sum(tabla_calculos['Variacio Setmanal'].head(c+1)>0)==c+1:
            c+=1
        index_metrics['Cotitzacions Consecutives setmanals'] = [c]
    else:
        c = 0
        while sum(tabla_calculos['Variacio Setmanal'].head(c+1)<0)==c+1:
            c+=1
        index_metrics['Cotitzacions Consecutives setmanals'] = [-c]

    # Cotitzacions mensuals consecutives
    if months['Diff%'][0]>0:
        c = 0
        while sum(months['Diff%'].head(c+1)>0)==c+1:
            c+=1
        index_metrics['Cotitzacions Consecutives mensuals'] = [c]
    else:
        c = 0
        while sum(months['Diff%'].head(c+1)<0)==c+1:
            c+=1
        index_metrics['Cotitzacions Consecutives mensuals'] = [-c]
    
    # RSI 10 Setmanes
    setmanes_10 = tabla_calculos['Variacio Setmanal'].head(10)
    RSI10pos = (setmanes_10[setmanes_10<0].count()/10)*100
    RSI10 = str(round(RSI10pos,2))+'/'+str(round(100-RSI10pos,2))
    index_metrics['RSI 10 Setmanes'] = [RSI10]

    # RSI Any en curs
    setmanes_any = tabla_calculos.loc[str(today.year), 'Variacio Setmanal']
    RSIany_pos = (setmanes_any[setmanes_any<0].count()/setmanes_any.count())*100
    RSIany = str(round(RSIany_pos,2))+' / '+str(round(100-RSIany_pos,2))
    index_metrics['RSI Any Actual'] = [RSIany]

    # Revalorització 1/5 al 9/10 
    if today > pd.Timestamp(date(year=today.year, month=5, day=1)).date() and today < pd.Timestamp(date(year=today.year, month=10, day=9)).date():
        abril = stock_data.loc[str(today.year)+'-04']
        inicio_p = abril.Close[0]
        dif_periodo = (stock_data.Close[0]/inicio_p-1)*100
        index_metrics['Revaloritzacio 1/5 - 9/10'] = [round(dif_periodo,2)]
    # Revalorització 1/10 al 30/04
    else:
        inicio_p = stock_data[stock_data.Month == 9]['Close'][0]
        dif_periodo = (stock_data.Close[0]/inicio_p-1)*100
        index_metrics['Revaloritzacio 1/10 - 30/04'] = [round(dif_periodo,2)]
    return index_metrics


def CreateDashboard(inds, stock_dict):
    cuadro_control = pd.DataFrame()
    for ind in inds:
        df = controlIndex(ind)
        df.index = [stock_dict[ind]]
        df = df.T
        cuadro_control = pd.concat([cuadro_control,df], axis=1)
    cuadro_control.reset_index(inplace=True)
    cuadro_control.rename(columns={'index':'Indicador'}, inplace=True)
    return cuadro_control

def trim_data(data):
    """
    Takes stock data
    Returns table with all the trimesters represented
    by the last day of the trimester
    """
    data_f = data.copy()
    data_f['Trim'] = data_f.Month.apply(lambda m: math.floor((m-1)/3)+1)
    data_f['LDT'] = data_f.Trim != data_f.Trim.shift(1)
    data_f = data_f[data_f['LDT']==True]
    data_f['Diff%'] = 100*(data_f.Close/data_f.Close.shift(-1)-1)
    return data_f

def years_data(data):
    """
    Takes stock data
    Returns table with all the years represented
    by the last day of the year
    """
    data_f = data.copy()
    data_f['LDY'] = data_f.Year != data_f.Year.shift(1)
    data_f = data_f[data_f['LDY']==True]
    data_f['Diff%'] = 100*(data_f.Close/data_f.Close.shift(-1)-1)
    return data_f

def year_trim_table(data):
    # Tabla central con suma por trimestres
    table_trim = year_month_table(data)
    table_trim['Trim 1'] = table_trim.iloc[:,:3].sum(axis = 1)
    table_trim['Trim 2'] = table_trim.iloc[:,3:6].sum(axis = 1)
    table_trim['Trim 3'] = table_trim.iloc[:,6:9].sum(axis = 1)
    table_trim['Trim 4'] = table_trim.iloc[:,9:12].sum(axis = 1)
    cols_order = ['Enero','Febrero','Marzo','Trim 1',
                 'Abril','Mayo','Junio','Trim 2',
                 'Julio','Agosto','Septiembre','Trim 3',
                 'Octubre','Noviembre','Diciembre','Trim 4']
    table_trim = table_trim[cols_order]

    return table_trim

def year_periods(data):
    table_ym = year_month_table(data)
    years = list(table_ym.index[::-1])
    years.append(years[-1]+1)
    index = [str(years[i])+'/'+str(years[i+1]) for i in range(len(years)-1)][::-1]
    # Periodo 1
    t_periodo1 = table_ym.copy()
    months = [ 'Noviembre', 'Diciembre', 'Enero', 'Febrero', 'Marzo', 'Abril']
    t_periodo1 = t_periodo1[months]
    t_periodo1[['Enero', 'Febrero', 'Marzo', 'Abril']] = t_periodo1[['Enero', 'Febrero', 'Marzo', 'Abril']].shift(1)
    t_periodo1.index = index
    t_periodo1['Total'] = t_periodo1.sum(axis=1)
    # Periodo 2
    t_periodo2 = table_ym.copy()
    months = ['Octubre', 'Noviembre', 'Diciembre', 'Enero', 'Febrero', 'Marzo']
    t_periodo2 = t_periodo2[months]
    t_periodo2[['Enero', 'Febrero', 'Marzo']] = t_periodo2[['Enero', 'Febrero', 'Marzo']].shift(1)
    t_periodo2.index = index
    t_periodo2['Total'] = t_periodo2.sum(axis=1)
    # Periodo 3
    months = ['Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre']
    t_periodo3 = table_ym.copy()
    t_periodo3 = t_periodo3[months]
    t_periodo3['Total'] = t_periodo3.sum(axis=1)
    #del t_periodo3.index.name
    # Promedios
    promedios = {'%Sube' : [100*t_periodo1.Total[t_periodo1.Total>0].count()/t_periodo1.Total[t_periodo1.Total!=0].count(),
                           100*t_periodo2.Total[t_periodo2.Total>0].count()/t_periodo2.Total[t_periodo2.Total!=0].count(),
                           100*t_periodo3.Total[t_periodo3.Total>0].count()/t_periodo3.Total[t_periodo3.Total!=0].count()],
                '%Baja' : [100*t_periodo1.Total[t_periodo1.Total<0].count()/t_periodo1.Total[t_periodo1.Total!=0].count(),
                           100*t_periodo2.Total[t_periodo2.Total<0].count()/t_periodo2.Total[t_periodo2.Total!=0].count(),
                           100*t_periodo3.Total[t_periodo3.Total<0].count()/t_periodo3.Total[t_periodo3.Total!=0].count()],
                'Subida' : [t_periodo1.Total[t_periodo1.Total>0].mean(),
                            t_periodo2.Total[t_periodo2.Total>0].mean(),
                            t_periodo3.Total[t_periodo3.Total>0].mean()],
                'Bajada' : [t_periodo1.Total[t_periodo1.Total<0].mean(),
                            t_periodo2.Total[t_periodo2.Total<0].mean(),
                            t_periodo3.Total[t_periodo3.Total<0].mean()]}
    promedios = pd.DataFrame(promedios, index = ['P1', 'P2', 'P3'])
    promedios = promedios.T
    return t_periodo1, t_periodo2, t_periodo3, promedios

def year_trim_aux(table):
    # Promedios trimestres
    promedios_trim = pd.DataFrame()
    trimestres = table[['Trim 1','Trim 2','Trim 3','Trim 4']]
    year_total = trimestres.sum(axis=1)
    year_total.header = 'Año'
    promedios_trim['%Sube'] = trimestres[trimestres>0].count(axis = 0)/trimestres.count(axis = 0)*100
    promedios_trim['%Baja'] = trimestres[trimestres<0].count(axis = 0)/trimestres.count(axis = 0)*100
    promedios_trim['Promedio Subida'] = trimestres[trimestres>0].mean(axis = 0)
    promedios_trim['Promedio Bajada'] = trimestres[trimestres<0].mean(axis = 0)
    promedios_trim = promedios_trim.T
    promedios_trim['Año'] = [len(year_total[year_total>0])/len(year_total)*100,
                            len(year_total[year_total<0])/len(year_total)*100,
                            year_total[year_total>0].mean(axis = 0),
                            year_total[year_total<0].mean(axis = 0)]
    # Trimestres alza / Trimestres baja
    alza = trimestres[trimestres>0].count(axis=1)
    baja = trimestres[trimestres<0].count(axis=1)
    alza_baja = pd.DataFrame({'Meses Alza':alza, 'Meses Baja':baja})
    # Tabla alza baja porcentaje
    trim_porcentajes = pd.DataFrame({'Alza/Baja':['0/4','1/3','2/2','3/1','4/0'],
                                     'Count':[0]*5})
    for i in range(alza_baja.shape[0]):
        if alza_baja.iloc[i,0]+alza_baja.iloc[i,1] == 4:
            trim_porcentajes.loc[alza_baja.iloc[i,0], 'Count'] += 1
    trim_porcentajes['Porcentaje'] = trim_porcentajes['Count']/trim_porcentajes['Count'].sum()*100
    trim_porcentajes.drop(columns='Count', inplace=True)
    return promedios_trim, year_total, alza_baja, trim_porcentajes

def notas_tables(data_i):
    data = data_i.copy()
    data['FDM'] = data.Month!=data.Month.shift(1) #Last Day of Month
    data['PDM'] = data.Month!=data.Month.shift(-1) #First Day of Month
    data['Close_ant'] = data.Close.shift(-1)
    years = list(data.Year.unique())[:3]  # Get the three most recent years
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
        df.index = ['Gener','Febrer','Març','Abril','Maig','Juny','Juliol','Agost','Septembre',
                   'Octubre','Novembre','Decembre'][:len(df.index)]
        notas_tables[year] = df
    return notas_tables

def read_data(ind): # Checked
    """
    Read data using yahoo finance library
    Extracts data from index = ind
    Adds column indicating day of the week
    Adds column for year, month, and day
    """
    try:
        if USE_CURL_CFFI:
            # Create a session that impersonates Chrome to avoid rate limits
            session = requests.Session(impersonate="chrome")
            index_ticker = yf.Ticker(ind, session=session)
        else:
            index_ticker = yf.Ticker(ind)
            
        # Add delay to avoid rate limiting
        time.sleep(0.2)
        
        index_data = pd.DataFrame(index_ticker.history(period="max"))
        index_data['Date'] = index_data.index
        index_data.index.rename('Date_index', inplace=True)
        weekday_map = {0:'Monday',1:'Tuesday',2:'Wednesday',3:'Thursday',4:'Friday'}
        #index_data['WeekDay'] = index_data.Date.dt.dayofweek.map(weekday_map) # Add WeekDay column
        index_data['WeekDay'] = index_data.Date.dt.dayofweek
        index_data['Week'] = index_data.Date.dt.isocalendar().week
        index_data['Year'] = index_data.Date.dt.year
        index_data['Month'] = index_data.Date.dt.month
        index_data['Day'] = index_data.Date.dt.day
        index_data = index_data[::-1]
        index_data['Diff%'] = 100*(index_data.Close/index_data.Close.shift(-1)-1)
        #index_data.Date = index_data.Date.astype(str)
        index_data.dropna(inplace=True)
        index_data.drop(columns = ['Dividends','Stock Splits'], inplace = True)
        return index_data
    except Exception as e:
        print(f"Error fetching data for {ind}: {e}")
        raise

def read_data_csv(index_data): # Checked
    """
    Read data using yahoo finance library
    Extracts data from index = ind
    Adds column indicating day of the week
    Adds column for year, month, and day
    """
    index_data['Date'] = index_data.index
    weekday_map = {0:'Monday',1:'Tuesday',2:'Wednesday',3:'Thursday',4:'Friday'}
    #index_data['WeekDay'] = index_data.Date.dt.dayofweek.map(weekday_map) # Add WeekDay column
    index_data['WeekDay'] = index_data.Date.dt.dayofweek
    index_data['Week'] = index_data.Date.dt.week
    index_data['Year'] = index_data.Date.dt.year
    index_data['Month'] = index_data.Date.dt.month
    index_data['Day'] = index_data.Date.dt.day
    index_data = index_data[::-1]
    index_data['Diff%'] = 100*(index_data.Close/index_data.Close.shift(-1)-1)
    #index_data.Date = index_data.Date.astype(str)
    index_data.dropna(inplace=True)
    index_data.drop(columns = ['Date'], inplace = True)
    return index_data

def months_data(data):
    """
    Takes stock data
    Returns table with all the months represented
    by the last day of the month
    """
    data_f = data.copy()
    data_f['LDM'] = data_f.Month != data_f.Month.shift(1)
    data_f = data_f[data_f['LDM']==True]
    data_f['Diff%'] = 100*(data_f.Close/data_f.Close.shift(-1)-1)
    return data_f

def year_month_table(data):
    """
    Takes stock data
    Returns a table with index = Year and columns = Month
    with the percentage changes respect to previous month
    """
    data_f = data.copy()
    data_f['LDM'] = data_f.Month != data_f.Month.shift(1)
    data_f = data_f[data_f['LDM']==True]
    data_f.drop(columns = ['LDM'], inplace=True)
    data_f['Diff%'] = 100*(data_f.Close/data_f.Close.shift(-1)-1)
    data_f = pd.pivot_table(data_f, values='Diff%', index='Year', columns='Month').\
            sort_values(by='Year', ascending=False)
    months_map = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
                 7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}
    data_f.rename(columns=months_map, inplace=True)
    return data_f

def year_month_aux(data):
    # Total año de tabla principal
    total_year = data.sum(axis=1)
    # Meses alza / Meses baja
    alza = data[data>0].count(axis=1)
    baja = data[data<0].count(axis=1)
    alza_baja = pd.DataFrame({'Meses Alza':alza, 'Meses Baja':baja})
    # Tabla medias de cada mes
    table_ym_1 = data.copy()
    table_ym_1 = table_ym_1.iloc[1:]
    media_t = table_ym_1.mean(axis=0, skipna=True)
    media_10 = table_ym_1.head(10).mean(axis=0, skipna=True)
    media_11 = table_ym_1.head(11).mean(axis=0, skipna=True)
    l = [np.abs(media_10*2), media_t, media_11, media_10]
    medias = pd.concat(l, axis=1).T
    medias = medias.rename(index={0:'Objetivo Actual', 1:'Media Total', 2:'Media ultimos 11', 3:'Media ultimos 10'})
    medias['Año'] = medias.sum(axis=1)
    # tabla promedios
    promedios = pd.DataFrame()
    promedios['%Sube'] = data[data>0].count(axis = 0)/data.count(axis = 0)*100
    promedios['%Baja'] = data[data<0].count(axis = 0)/data.count(axis = 0)*100
    promedios['Promedio Subida'] = data[data>0].mean(axis = 0)
    promedios['Promedio Bajada'] = data[data<0].mean(axis = 0)
    promedios = promedios.T
    promedios['Año'] = [total_year[total_year>0].count()/total_year.count()*100,
                        total_year[total_year<0].count()/total_year.count()*100,
                        total_year[total_year>0].mean(),
                        total_year[total_year<0].mean()]
    # tabla alza baja porcentajes
    meses_porcentajes = pd.DataFrame({'Frac':['0/12','1/11','2/10','3/9','4/8','5/7','6/6','7/6',
                    '8/4','9/3','10/2','11/1','12/0'],
                                     'Count':[0]*13})
    for i in range(alza_baja.shape[0]):
        if alza_baja.iloc[i,0]+alza_baja.iloc[i,1] == 12:
            meses_porcentajes.loc[alza_baja.iloc[i,0], 'Count'] += 1
    meses_porcentajes['Perc'] = meses_porcentajes['Count']/meses_porcentajes['Count'].sum()*100
    meses_porcentajes.drop(columns='Count', inplace=True)
    return total_year, alza_baja, medias, promedios, meses_porcentajes

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

def alarms_baja(stock_data):
    alarms = pd.DataFrame(columns=['Alarma Baixada', 'Valor','Estat'])
    today = stock_data.index[0]
    today_date = today.date() # Added for date comparison

    # Alarma 1: Bajada de 9.15% entre 1 de mayo y 9 de octubre
    alarm_dsc = 'Baixada del periode superior al 9.15% (01/05 - 09/10)'
    if today_date > date(year=today.year,month=5, day=1) and today_date < date(year=today.year,month=10, day=9):
        abril = stock_data.loc[str(today.year)+'-04']
        inicio_p = abril.Close[0]
        dif_periodo = (stock_data.Close[0]/inicio_p-1)*100
        if dif_periodo < -9.15:
            new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(dif_periodo,2)],'Estat':pd.Series([True], dtype='boolean')})
            alarms = pd.concat([alarms, new_row], ignore_index=True)
        else:
            new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(dif_periodo,2)],'Estat':pd.Series([False], dtype='boolean')})
            alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor': [np.nan], 'Estat':pd.Series([pd.NA], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 2: Bajada semanal del 2.56% en el periodo del 01/05 al 09/10
    #           Bajada semanal del 2.19% en el periodo del 10/10 al 30/04
    tabla_calculos, promedios_calc = calculos_table(stock_data)

    if today_date > date(year=today.year, month=5, day=1) and today_date < date(year=today.year, month=10, day=9):
        alarm_dsc = 'Baixada setmanal superior al 2.56% (01/05 - 09/10)'
        threshold = -2.56
    else:
        alarm_dsc = 'Baixada setmanal superior al 2.19% (10/10 - 30/04)'
        threshold = -2.19

    dif_set = tabla_calculos['Variacio Setmanal'][0]
    if dif_set < threshold:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(dif_set, 2)], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(dif_set, 2)], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 3: Cuando de las ultimas 10 semanas, el indice ha bajado 8
    var_x_10 = sum(tabla_calculos['Variacio Setmanal'][0:10]<0)
    alarm_dsc = 'De les ultimes 10 setmanes, l\'index ha baixat 8 o mes cops'
    if var_x_10>=8:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[var_x_10], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[var_x_10], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 4: Cuando lleva 5 lunes seguidos bajando
    lun, mar, mie, jue, vie, promedios = tablas_diario(stock_data)
    alarm_dsc = 'Porta 5 dilluns seguits baixant'
    c = 0
    while sum(lun['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 5: Cuando lleva 5 martes seguidos bajando
    alarm_dsc = 'Porta 5 dimarts seguits baixant'
    c = 0
    while sum(mar['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 6: Cuando lleva 5 miercoles seguidos bajando
    alarm_dsc = 'Porta 5 dimecres seguits baixant'
    c = 0
    while sum(mie['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 7: Cuando lleva 5 jueves seguidos bajando
    alarm_dsc = 'Porta 5 dijous seguits baixant'
    c = 0
    while sum(jue['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 8: Cuando lleva 5 viernes seguidos bajando
    alarm_dsc = 'Porta 5 divendres seguits baixant'
    c = 0
    while sum(vie['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 9: Cuando lleva 5 dias seguidos bajando
    alarm_dsc = 'Porta 5 dies seguits baixant'
    c = 0
    while sum(stock_data['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 10: Cuando lleva 4 semanas seguidas bajando
    alarm_dsc = 'Porta 4 setmanes seguides baixant'
    c = 0
    while sum(tabla_calculos['Variacio Setmanal'].head(c+1)<0)==c+1:
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 11: Cuando lleva 4 meses seguidos bajando
    alarm_dsc = 'Porta 4 mesos seguits baixant'
    months = months_data(stock_data)
    c = 0
    while sum(months['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 12: Cuando lleva 4 trimestres seguidos bajando
    alarm_dsc = 'Porta 4 trimestres seguits baixant'
    trims = trim_data(stock_data)
    c = 0
    while sum(trims['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 13: Cuando lleva 4 años seguidos bajando
    alarm_dsc = 'Porta 4 anys seguits baixant'
    years_df = years_data(stock_data) # Renamed to avoid conflict with 'years' variable in generarExcel
    c = 0
    while sum(years_df['Diff%'].head(c+1)<0)==c+1: # Use years_df here
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 14: Cuando el mes actual ha superado su media a la baja
    alarm_dsc = 'El mes actual ha superat la seva mitja a la baixa'
    table_ym = year_month_table(stock_data)
    total_year_ym, alza_baja_ym, medias_ym, promedios_ym, meses_porcentajes_ym = year_month_aux(table_ym)
    if months['Diff%'][0]<promedios_ym.loc['Promedio Bajada'][today_date.month-1]: # Use today_date here
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(months['Diff%'][0],2)], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(months['Diff%'][0],2)], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 15: Cuando la semana actual ha superado el porcentaje de variacion de las ultimas 10 semanas
    alarm_dsc = 'La setmana actual ha superat la mitja de variació de les últimes 10 setmanes a la baixa'
    if tabla_calculos['Variacio Setmanal'][0]<-tabla_calculos['Porc. Var. 10'][0]:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(tabla_calculos['Variacio Setmanal'][0],2)], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(tabla_calculos['Variacio Setmanal'][0],2)], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 16: Cuando el 4º viernes del mes acaba en negativo
    alarm_dsc = 'L\'ultim quart divendres de mes va acabar en negatiu'
    vie100 = pd.DataFrame(pd.date_range(end = today_date, periods=100), columns = ['Date'])[::-1] # Use today_date here
    vie100.Date = pd.to_datetime(vie100.Date)
    vie100 = vie100[vie100.Date.dt.weekday == 4]
    vie100['4FM'] = (vie100.Date.dt.month.shift(-4)!=vie100.Date.dt.month) & \
                    (vie100.Date.dt.month.shift(-3)==vie100.Date.dt.month)
    vie100 = vie100[vie100['4FM']==True]

    # Prepare stock_data for merge by ensuring its 'Date' column is timezone-naive
    stock_data_for_merge = stock_data.copy()
    if 'Date' in stock_data_for_merge.columns and \
       pd.api.types.is_datetime64_any_dtype(stock_data_for_merge['Date']) and \
       getattr(stock_data_for_merge['Date'].dt, 'tz', None) is not None:
        stock_data_for_merge['Date'] = stock_data_for_merge['Date'].dt.tz_localize(None)
    
    # Ensure vie100['Date'] is also naive (it should be by construction, but this is a defensive check)
    if 'Date' in vie100.columns and \
       pd.api.types.is_datetime64_any_dtype(vie100['Date']) and \
       getattr(vie100['Date'].dt, 'tz', None) is not None:
        vie100['Date'] = vie100['Date'].dt.tz_localize(None)
        
    vie100 = pd.merge(vie100, stock_data_for_merge, on = 'Date', how='left') # Use the modified stock_data and how='left'
    if not vie100.empty and vie100['Diff%'][0]<0: # Added check for empty vie100
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(vie100['Diff%'][0],2)], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    elif not vie100.empty: # Added check for empty vie100
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[round(vie100['Diff%'][0],2)], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else: # Handle case where vie100 is empty
        new_row = pd.DataFrame({'Alarma Baixada':[alarm_dsc], 'Valor':[np.nan], 'Estat':pd.Series([pd.NA], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Cast 'Estat' column to bool before reductions
    alarms['Estat'] = alarms['Estat'].astype(bool)
    return alarms
    
def alarms_alza(stock_data):
    alarms = pd.DataFrame(columns=['Alarma Pujada', 'Valor','Estat'])
    today = stock_data.index[0]
    today_date = today.date() # Added for date comparison
 
    # Alarma 1: Cuando de las ultimas 10 semanas, el indice ha bajado 8
    tabla_calculos, promedios_calc = calculos_table(stock_data)
    var_x_10 = sum(tabla_calculos['Variacio Setmanal'][0:10]>0)
    alarm_dsc = 'De les ultimes 10 setmanes, l\'index ha pujat 8 o mes cops'
    if var_x_10>=8:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[var_x_10], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[var_x_10], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 2: Cuando lleva 5 lunes seguidos subiendo
    lun, mar, mie, jue, vie, promedios = tablas_diario(stock_data)
    alarm_dsc = 'Porta 5 dilluns seguits pujant'
    c = 0
    while sum(lun['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 3: Cuando lleva 5 martes seguidos subiendo
    alarm_dsc = 'Porta 5 dimarts seguits pujant'
    c = 0
    while sum(mar['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 4: Cuando lleva 5 miercoles seguidos subiendo
    alarm_dsc = 'Porta 5 dimecres seguits pujant'
    c = 0
    while sum(mie['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 5: Cuando lleva 5 jueves seguidos subiendo
    alarm_dsc = 'Porta 5 dijous seguits pujant'
    c = 0
    while sum(jue['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 6: Cuando lleva 5 viernes seguidos subiendo
    alarm_dsc = 'Porta 5 divendres seguits pujant'
    c = 0
    while sum(vie['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 7: Cuando lleva 5 dias seguidos subiendo
    alarm_dsc = 'Porta 5 dies seguits pujant'
    c = 0
    while sum(stock_data['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 8: Cuando lleva 4 semanas seguidas subiendo
    alarm_dsc = 'Porta 4 setmanes seguides pujant'
    c = 0
    while sum(tabla_calculos['Variacio Setmanal'].head(c+1)>0)==c+1:
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 9: Cuando lleva 4 meses seguidos subiendo
    alarm_dsc = 'Porta 4 mesos seguits pujant'
    months = months_data(stock_data)
    c = 0
    while sum(months['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 10: Cuando lleva 4 trimestres seguidos subiendo
    alarm_dsc = 'Porta 4 trimestres seguits pujant'
    trims = trim_data(stock_data)
    c = 0
    while sum(trims['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)

    # Alarma 11: Cuando lleva 4 años seguidos subiendo
    alarm_dsc = 'Porta 4 anys seguits pujant'
    years_df = years_data(stock_data) # Renamed to avoid conflict
    c = 0
    while sum(years_df['Diff%'].head(c+1)>0)==c+1: # Use years_df here
        c+=1
    if c>=4:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[c], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 12: Cuando el mes actual ha superado su media a la alza
    alarm_dsc = 'El mes actual ha superat la seva mitja a la alça'
    table_ym = year_month_table(stock_data)
    total_year_ym, alza_baja_ym, medias_ym, promedios_ym, meses_porcentajes_ym = year_month_aux(table_ym)
    if months['Diff%'][0]>promedios_ym.loc['Promedio Subida'][today_date.month-1]: # Use today_date here
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[round(months['Diff%'][0],2)], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[round(months['Diff%'][0],2)], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Alarma 13: Cuando la semana actual ha superado el porcentaje de variacion de las ultimas 10 semanas
    alarm_dsc = 'La setmana actual ha superat la mitja de variació de les últimes 10 setmanes a la alça'
    if tabla_calculos['Variacio Setmanal'][0]>tabla_calculos['Porc. Var. 10'][0]:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[round(tabla_calculos['Variacio Setmanal'][0],2)], 'Estat':pd.Series([True], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    else:
        new_row = pd.DataFrame({'Alarma Pujada':[alarm_dsc], 'Valor':[round(tabla_calculos['Variacio Setmanal'][0],2)], 'Estat':pd.Series([False], dtype='boolean')})
        alarms = pd.concat([alarms, new_row], ignore_index=True)
    
    # Cast 'Estat' column to bool before reductions
    alarms['Estat'] = alarms['Estat'].astype(bool)
    return alarms

def generarExcel(ind):
    stock_data = read_data(ind)

    today = date.today()
    filename = 'Informe'+stock_dict[ind]+str(today.year)+'-'+str(today.month)+'-'+str(today.day)+'.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook  = writer.book

    # Formats
    red_text = workbook.add_format({'font_color': 'red', 'num_format': '#,##0.00'})
    red_bold = workbook.add_format({'font_color': 'red', 'bold':True})
    green_bold = workbook.add_format({'font_color': 'green', 'bold':True})
    black_text = workbook.add_format({'font_color': 'black', 'num_format': '#,##0.00'})
    bold_str = workbook.add_format({'bold':True})
    bold = workbook.add_format({'bold':True, 'num_format': '#,##0.00'})
    centered = workbook.add_format()
    centered.set_align('center')
    centered.set_align('vcenter')

    # Cuadro Control
    dashboard = CreateDashboard(inds, stock_dict)
    dashboard.to_excel(writer, sheet_name='Quadre Control', startrow = 1, startcol = 1, index=False)

    worksheet = writer.sheets['Quadre Control']
    worksheet.set_column('B:B', 70)
    # Dynamically format columns based on the number of indices
    for i in range(len(inds) + 1):
        col = chr(ord('C') + i)
        worksheet.set_column(f'{col}:{col}', 20, centered)

    #Alarmas
    alarmas1 = alarms_baja(stock_data)
    alarmas2 = alarms_alza(stock_data)
    alarmas1.index = alarmas1.index+1
    alarmas2.index = alarmas2.index+1
    alarmas1.to_excel(writer, sheet_name = 'Alarmas', startrow = 1, startcol = 1)
    alarmas2.to_excel(writer, sheet_name = 'Alarmas', startrow = 1, startcol = 6)

    worksheet = writer.sheets['Alarmas']
    worksheet.set_column('C:C', 70)
    worksheet.set_column('H:H', 70)
    worksheet.set_column('E:E', 15)
    worksheet.set_column('J:J', 15)
    worksheet.conditional_format('E3:E1000', {'type':     'cell',
                                            'criteria': '=',
                                            'value':    True,
                                            'format':   green_bold})
    worksheet.conditional_format('J3:J1000', {'type':     'cell',
                                            'criteria': '=',
                                            'value':    True,
                                            'format':   red_bold})


    # Resumen Mensual
    table_ym = year_month_table(stock_data)
    total_year_ym, alza_baja_ym, medias_ym, promedios_ym, meses_porcentajes_ym = year_month_aux(table_ym)

    medias_ym.to_excel(writer, sheet_name='Resumen Mensual', startcol=1, startrow=1)
    table_ym.to_excel(writer, sheet_name='Resumen Mensual', startcol=1, startrow=7, header=False)
    total_year_ym.to_excel(writer, sheet_name='Resumen Mensual', startcol=14, startrow=7, header=False, index=False)
    years = pd.DataFrame(table_ym.index)
    years.to_excel(writer, sheet_name='Resumen Mensual', startcol=15, startrow=7, header=False, index=False)
    promedios_ym.to_excel(writer, sheet_name='Resumen Mensual', startcol=1, startrow=8+table_ym.shape[0], header=False)
    alza_baja_ym.to_excel(writer, sheet_name='Resumen Mensual', startcol=17, startrow=6, index=False)
    meses_porcentajes_ym.to_excel(writer, sheet_name='Resumen Mensual', startcol=20, startrow=7, index=False, header=False)

    worksheet = writer.sheets['Resumen Mensual']
    worksheet.conditional_format('C3:O1000', {'type':     'cell',
                                            'criteria': '<',
                                            'value':    0,
                                            'format':   red_text})
    worksheet.conditional_format('C3:O1000', {'type':     'cell',
                                            'criteria': '>',
                                            'value':    0,
                                            'format':   black_text})
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:V', 10)
    worksheet.set_column('P:P', 10, bold_str)
    worksheet.set_column('U:V', 10, black_text)

    #Resumen Trimestral
    table_trim = year_trim_table(stock_data)
    promedios_trim, year_total, alza_baja, trim_porcentajes = year_trim_aux(table_trim)
    table_trim.to_excel(writer, sheet_name='Resumen Trimestral', startcol=1, startrow=1)
    year_total = pd.DataFrame(year_total,columns=['Año'])
    year_total.to_excel(writer, sheet_name='Resumen Trimestral', startcol=18, startrow=1, index=False)
    years = pd.DataFrame(table_trim.index)
    years.to_excel(writer, sheet_name='Resumen Trimestral', startcol=19, startrow=2, header=False, index=False)
    alza_baja.to_excel(writer, sheet_name='Resumen Trimestral', startcol=21, startrow=1, index=False)
    index_prom = pd.DataFrame(promedios_trim.index)
    index_prom.to_excel(writer, sheet_name='Resumen Trimestral', startcol=1, startrow=3+table_trim.shape[0],
                       header=False, index=False)
    promedios_trim.iloc[:,0].to_excel(writer, sheet_name='Resumen Trimestral', startcol=5, startrow=3+table_trim.shape[0],
                              header=False, index=False)
    promedios_trim.iloc[:,1].to_excel(writer, sheet_name='Resumen Trimestral', startcol=9, startrow=3+table_trim.shape[0],
                              header=False, index=False)
    promedios_trim.iloc[:,2].to_excel(writer, sheet_name='Resumen Trimestral', startcol=13, startrow=3+table_trim.shape[0],
                              header=False, index=False)
    promedios_trim.iloc[:,3].to_excel(writer, sheet_name='Resumen Trimestral', startcol=17, startrow=3+table_trim.shape[0],
                              header=False, index=False)
    trim_porcentajes.to_excel(writer, sheet_name='Resumen Trimestral', startcol=23, startrow=1,
                              index=False)

    worksheet = writer.sheets['Resumen Trimestral']
    worksheet.conditional_format('C3:S1000', {'type':     'cell',
                                            'criteria': '<',
                                            'value':    0,
                                            'format':   red_text})
    worksheet.conditional_format('C3:S1000', {'type':     'cell',
                                            'criteria': '>',
                                            'value':    0,
                                            'format':   black_text})
    worksheet.set_column('B:W', 10)
    worksheet.set_column('F1:F100', 10, bold)
    worksheet.set_column('J1:J100', 10, bold)
    worksheet.set_column('N1:N100', 10, bold)
    worksheet.set_column('R1:R100', 10, bold)
    worksheet.set_column('X1:Y100', 10, black_text)

    # Periodos que estaban en trimestres
    t_periodo1, t_periodo2, t_periodo3, promedios_p = year_periods(stock_data)
    t_periodo1.to_excel(writer, sheet_name='Periodos', startcol=1, startrow=1)
    t_periodo2.to_excel(writer, sheet_name='Periodos', startcol=10, startrow=1)
    t_periodo3.to_excel(writer, sheet_name='Periodos', startcol=19, startrow=1)
    index_prom = pd.DataFrame(promedios_p.index)
    index_prom.to_excel(writer, sheet_name='Periodos', startcol=1, startrow=3+t_periodo1.shape[0],
                       header=False, index=False)
    index_prom.to_excel(writer, sheet_name='Periodos', startcol=10, startrow=3+t_periodo2.shape[0],
                       header=False, index=False)
    index_prom.to_excel(writer, sheet_name='Periodos', startcol=19, startrow=3+t_periodo3.shape[0],
                       header=False, index=False)
    promedios_p.iloc[:,0].to_excel(writer, sheet_name='Periodos', startcol=8, startrow=3+t_periodo1.shape[0],
                              header=False, index=False)
    promedios_p.iloc[:,1].to_excel(writer, sheet_name='Periodos', startcol=17, startrow=3+t_periodo2.shape[0],
                              header=False, index=False)
    promedios_p.iloc[:,2].to_excel(writer, sheet_name='Periodos', startcol=26, startrow=3+t_periodo3.shape[0],
                              header=False, index=False)

    worksheet = writer.sheets['Periodos']
    worksheet.conditional_format('C3:R1000', {'type':     'cell',
                                            'criteria': '<',
                                            'value':    0,
                                            'format':   red_text})
    worksheet.conditional_format('C3:R1000', {'type':     'cell',
                                            'criteria': '>',
                                            'value':    0,
                                            'format':   black_text})
    worksheet.conditional_format('U3:AA1000', {'type':     'cell',
                                            'criteria': '<',
                                            'value':    0,
                                            'format':   red_text})
    worksheet.conditional_format('U3:AA1000', {'type':     'cell',
                                            'criteria': '>',
                                            'value':    0,
                                            'format':   black_text})
    worksheet.set_column('C1:AA100', 10)
    worksheet.set_column('I:I', 10, bold)
    worksheet.set_column('R:R', 10, bold)
    worksheet.set_column('AA:AA', 10, bold)
    worksheet.set_column('B31:B35', 10, bold)
    worksheet.set_column('K:K', 10, bold)
    worksheet.set_column('T:T', 10, bold_str)

    # Calculos
    #semanas=20
    tabla_calculos, promedios_calc = calculos_table(stock_data)
    tabla_calculos.index = tabla_calculos.index.strftime('%d/%m/%Y')
    tabla_calculos.to_excel(writer, sheet_name='Calculos', startcol=1, startrow=1)
    promedios_calc.to_excel(writer, sheet_name='Calculos', startcol=3+tabla_calculos.shape[1], startrow=1)
    worksheet = writer.sheets['Calculos']
    worksheet.set_column('B1:Z100', 15, black_text)
    worksheet.set_column('L1:L100', 15, bold)
    worksheet.conditional_format('C1:C100', {'type':     'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   red_text})

    # Notas
    notas_tabs = notas_tables(stock_data)
    years = list(notas_tabs.keys())
    
    # Add spacing and headers for each year
    row_position = 1
    for i, year in enumerate(years):
        notas_tabs[year].to_excel(writer, sheet_name='Notas', startrow=row_position, startcol=1)
        worksheet = writer.sheets['Notas']
        worksheet.write(row_position, 1, str(year))
        row_position += 15  # Add spacing between tables
    
    worksheet.set_column('B1:G100', 10, black_text)
    worksheet.conditional_format('E1:E100', {'type':     'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   red_text})
    worksheet.conditional_format('G1:G100', {'type':     'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   red_text})
    # Diario
    lun, mar, mie, jue, vie, promedios = tablas_diario(stock_data)
    header = pd.DataFrame(lun.columns[:-1]).T

    day_data = lun.iloc[:,:-1]
    day = pd.DataFrame(lun.iloc[:,-1])
    prom = pd.DataFrame(promedios[0])
    header.to_excel(writer, sheet_name='Resumen Diario', startrow=0, startcol=3, header=False, index=False)
    day_data.to_excel(writer, sheet_name = 'Resumen Diario', startrow = 8, startcol = 2, header = False)
    prom.to_excel(writer, sheet_name='Resumen Diario', startrow=2, startcol=4, header=False)
    day.to_excel(writer, sheet_name='Resumen Diario', startrow=8, startcol=1, header=False, index=False)

    day_data = mar.iloc[:,:-1]
    day = pd.DataFrame(mar.iloc[:,-1])
    prom = pd.DataFrame(promedios[1])
    header.to_excel(writer, sheet_name='Resumen Diario', startrow=0, startcol=12, header=False, index=False)
    day_data.to_excel(writer, sheet_name = 'Resumen Diario', startrow = 8, startcol = 11, header = False)
    prom.to_excel(writer, sheet_name='Resumen Diario', startrow=2, startcol=13, header=False)
    day.to_excel(writer, sheet_name='Resumen Diario', startrow=8, startcol=10, header=False, index=False)

    day_data = mie.iloc[:,:-1]
    day = pd.DataFrame(mie.iloc[:,-1])
    prom = pd.DataFrame(promedios[2])
    header.to_excel(writer, sheet_name='Resumen Diario', startrow=0, startcol=21, header=False, index=False)
    day_data.to_excel(writer, sheet_name = 'Resumen Diario', startrow = 8, startcol = 20, header = False)
    prom.to_excel(writer, sheet_name='Resumen Diario', startrow=2, startcol=22, header=False)
    day.to_excel(writer, sheet_name='Resumen Diario', startrow=8, startcol=19, header=False, index=False)

    day_data = jue.iloc[:,:-1]
    day = pd.DataFrame(jue.iloc[:,-1])
    prom = pd.DataFrame(promedios[3])
    header.to_excel(writer, sheet_name='Resumen Diario', startrow=0, startcol=30, header=False, index=False)
    day_data.to_excel(writer, sheet_name = 'Resumen Diario', startrow = 8, startcol = 29, header = False)
    prom.to_excel(writer, sheet_name='Resumen Diario', startrow=2, startcol=31, header=False)
    day.to_excel(writer, sheet_name='Resumen Diario', startrow=8, startcol=28, header=False, index=False)

    day_data = vie.iloc[:,:-1]
    day = pd.DataFrame(vie.iloc[:,-1])
    prom = pd.DataFrame(promedios[4])
    header.to_excel(writer, sheet_name='Resumen Diario', startrow=0, startcol=39, header=False, index=False)
    day_data.to_excel(writer, sheet_name = 'Resumen Diario', startrow = 8, startcol = 38, header = False)
    prom.to_excel(writer, sheet_name='Resumen Diario', startrow=2, startcol=40, header=False)
    day.to_excel(writer, sheet_name='Resumen Diario', startrow=8, startcol=37, header=False, index=False)

    worksheet = writer.sheets['Resumen Diario']
    worksheet.set_column('B1:H10000', 10, black_text)
    worksheet.set_column('I1:I10000', 12)
    worksheet.set_column('J1:Q10000', 10, black_text)
    worksheet.set_column('R1:R10000', 12)
    worksheet.set_column('S1:Z10000', 10, black_text)
    worksheet.set_column('AA1:AA10000', 12)
    worksheet.set_column('AB1:AI10000', 10, black_text)
    worksheet.set_column('AJ1:AJ10000', 12)
    worksheet.set_column('AK1:AR10000', 10, black_text)
    worksheet.set_column('AS1:AS10000', 12)
    worksheet.conditional_format('F1:F10000', {'type':     'cell',
                                               'criteria': '<',
                                               'value':    0,
                                               'format':   red_text})
    worksheet.conditional_format('O1:O10000', {'type':     'cell',
                                               'criteria': '<',
                                               'value':    0,
                                               'format':   red_text})
    worksheet.conditional_format('X1:X10000', {'type':     'cell',
                                               'criteria': '<',
                                               'value':    0,
                                               'format':   red_text})
    worksheet.conditional_format('AG1:AG10000', {'type':     'cell',
                                                 'criteria': '<',
                                                 'value':    0,
                                                 'format':   red_text})
    worksheet.conditional_format('AP1:AP10000', {'type':     'cell',
                                                 'criteria': '<',
                                                 'value':    0,
                                                 'format':   red_text})
    writer.close()
    
num = 1
while num!=0:
    num = int(input('Quin informe vols generar?\n1.IBEX\n2.NASDAQ\n3.DOWJONES\n4.SP500\n5.NASDAQ 100\n6.EURO STOXX 50\n7.DAX\n8.CAC 40\n9.FTSE 100\n10.HANG SENG\n0.EXIT\n--> '))
    # Read tables
    if num!=0:
        inds = ["^IBEX", "^IXIC", "^DJI", "^GSPC", "^NDX", "^STOXX50E", "^GDAXI", "^FCHI", "^FTSE", "^HSI"]
        ind = inds[num-1]
        stock_dict = {"^IBEX":"IBEX", "^IXIC":"NASDAQ", "^DJI":"DOWJONES", "^GSPC": "SP500", 
                     "^NDX": "NASDAQ100", "^STOXX50E": "EUROSTOXX50", "^GDAXI": "DAX", 
                     "^FCHI": "CAC40", "^FTSE": "FTSE100", "^HSI": "HANGSENG"}
        generarExcel(ind)
