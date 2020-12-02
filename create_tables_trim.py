import pandas as pd
import numpy as np
import yfinance as yf
import math
from create_tables_month import read_data, year_month_table, year_month_aux

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
            trim_porcentajes.iloc[alza_baja.iloc[i,0],1]+=1
    trim_porcentajes['Porcentaje'] = trim_porcentajes['Count']/trim_porcentajes['Count'].sum()*100
    trim_porcentajes.drop(columns='Count', inplace=True)
    return promedios_trim, year_total, alza_baja, trim_porcentajes
