from create_tables_month import read_data
import pandas as pd
from datetime import date
from create_table_calculos import calculos_table
from create_tables_month import months_data

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
    setmanes_any = tabla_calculos[str(today.year)]['Variacio Setmanal']
    RSIany_pos = (setmanes_any[setmanes_any<0].count()/setmanes_any.count())*100
    RSIany = str(round(RSIany_pos,2))+' / '+str(round(100-RSIany_pos,2))
    index_metrics['RSI Any Actual'] = [RSIany]

    # Revalorització 1/5 al 9/10 
    if today>date(year=today.year,month=5, day=1) and today<date(year=today.year,month=10, day=9):
        abril = stock_data[str(today.year)+'-04']
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