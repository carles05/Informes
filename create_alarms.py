import pandas as pd
import numpy as np
from create_tables_month import read_data
from datetime import date
from create_table_calculos import calculos_table
from create_resumen_diario import tablas_diario
from create_tables_month import months_data, year_month_aux, year_month_table
from create_tables_trim import trim_data, years_data


def alarms_baja(stock_data):
    alarms = pd.DataFrame(columns=['Alarma Baixada', 'Valor','Estat'])
    today = stock_data.index[0]

    # Alarma 1: Bajada de 9.15% entre 1 de mayo y 9 de octubre
    alarm_dsc = 'Baixada del periode superior al 9.15% (01/05 - 09/10)'
    if today>date(year=today.year,month=5, day=1) and today<date(year=today.year,month=10, day=9):
        abril = stock_data[str(today.year)+'-04']
        inicio_p = abril.Close[0]
        dif_periodo = (stock_data.Close[0]/inicio_p-1)*100
        if dif_periodo < -9.15:
            alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor':round(dif_periodo,2),'Estat':True}, ignore_index=True)
        else:
            alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor':round(dif_periodo,2),'Estat':False}, ignore_index=True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': np.nan, 'Estat':np.nan}, ignore_index=True)
    
    # Alarma 2: Bajada semanal del 2.56% en el periodo del 01/05 al 09/10
    #           Bajada semanal del 2.19% en el periodo del 10/10 al 30/04
    tabla_calculos, promedios_calc = calculos_table(stock_data)

    if today>date(year=today.year,month=5, day=1) and today<date(year=today.year,month=10, day=9):
        alarm_dsc = 'Baixada setmanal superior al 2.56% (01/05 - 09/10)'
        threshold = -2.56
    else:
        alarm_dsc = 'Baixada setmanal superior al 2.19% (10/10 - 30/04)'
        threshold = -2.19

    dif_set = tabla_calculos['Variacio Setmanal'][0]
    if dif_set < threshold:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor':round(dif_set, 2), 'Estat':True}, ignore_index=True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor':round(dif_set, 2), 'Estat':False}, ignore_index=True)
    
    # Alarma 3: Cuando de las ultimas 10 semanas, el indice ha bajado 8
    var_x_10 = sum(tabla_calculos['Variacio Setmanal'][0:10]<0)
    alarm_dsc = 'De les ultimes 10 setmanes, l\'index ha baixat 8 o mes cops'
    if var_x_10>=8:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor':var_x_10, 'Estat':True}, ignore_index=True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor':var_x_10, 'Estat':False}, ignore_index=True)

    # Alarma 4: Cuando lleva 5 lunes seguidos bajando
    lun, mar, mie, jue, vie, promedios = tablas_diario(stock_data)
    alarm_dsc = 'Porta 5 dilluns seguits baixant'
    c = 0
    while sum(lun['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)
    
    # Alarma 5: Cuando lleva 5 martes seguidos bajando
    alarm_dsc = 'Porta 5 dimarts seguits baixant'
    c = 0
    while sum(mar['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 6: Cuando lleva 5 miercoles seguidos bajando
    alarm_dsc = 'Porta 5 dimecres seguits baixant'
    c = 0
    while sum(mie['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 7: Cuando lleva 5 jueves seguidos bajando
    alarm_dsc = 'Porta 5 dijous seguits baixant'
    c = 0
    while sum(jue['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 8: Cuando lleva 5 viernes seguidos bajando
    alarm_dsc = 'Porta 5 divendres seguits baixant'
    c = 0
    while sum(vie['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)
    
    # Alarma 9: Cuando lleva 5 dias seguidos bajando
    alarm_dsc = 'Porta 5 dies seguits baixant'
    c = 0
    while sum(stock_data['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 10: Cuando lleva 4 semanas seguidas bajando
    alarm_dsc = 'Porta 4 setmanes seguides baixant'
    c = 0
    while sum(tabla_calculos['Variacio Setmanal'].head(c+1)<0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 11: Cuando lleva 4 meses seguidos bajando
    alarm_dsc = 'Porta 4 mesos seguits baixant'
    months = months_data(stock_data)
    c = 0
    while sum(months['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 12: Cuando lleva 4 trimestres seguidos bajando
    alarm_dsc = 'Porta 4 trimestres seguits baixant'
    trims = trim_data(stock_data)
    c = 0
    while sum(trims['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 13: Cuando lleva 4 años seguidos bajando
    alarm_dsc = 'Porta 4 anys seguits baixant'
    years = years_data(stock_data)
    c = 0
    while sum(years['Diff%'].head(c+1)<0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)
    
    # Alarma 14: Cuando el mes actual ha superado su media a la baja
    alarm_dsc = 'El mes actual ha superat la seva mitja a la baixa'
    table_ym = year_month_table(stock_data)
    total_year_ym, alza_baja_ym, medias_ym, promedios_ym, meses_porcentajes_ym = year_month_aux(table_ym)
    if months['Diff%'][0]<promedios_ym.loc['Promedio Bajada'][today.month-1]:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': round(months['Diff%'][0],2), 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': round(months['Diff%'][0],2), 'Estat': False}, ignore_index = True)
    
    # Alarma 15: Cuando la semana actual ha superado el porcentaje de variacion de las ultimas 10 semanas
    alarm_dsc = 'La setmana actual ha superat la mitja de variació de les últimes 10 setmanes a la baixa'
    if tabla_calculos['Variacio Setmanal'][0]<-tabla_calculos['Porc. Var. 10'][0]:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': round(tabla_calculos['Variacio Setmanal'][0],2),\
                                'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': round(tabla_calculos['Variacio Setmanal'][0],2),\
                                'Estat': False}, ignore_index = True)

    # Alarma 16: Cuando el 4º viernes del mes acaba en negativo
    alarm_dsc = 'L\'ultim quart divendres de mes va acabar en negatiu'
    vie100 = pd.DataFrame(pd.date_range(end = today, periods=100), columns = ['Date'])[::-1]
    vie100.Date = pd.to_datetime(vie100.Date)
    vie100 = vie100[vie100.Date.dt.weekday == 4]
    vie100['4FM'] = (vie100.Date.dt.month.shift(-4)!=vie100.Date.dt.month) &\
                    (vie100.Date.dt.month.shift(-3)==vie100.Date.dt.month)
    vie100 = vie100[vie100['4FM']==True]
    vie100 = pd.merge(vie100, stock_data, on = 'Date')
    if vie100['Diff%'][0]<0:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': round(vie100['Diff%'][0],2), 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Baixada':alarm_dsc, 'Valor': round(vie100['Diff%'][0],2), 'Estat': False}, ignore_index = True)
    return alarms
    
def alarms_alza(stock_data):
    alarms = pd.DataFrame(columns=['Alarma Pujada', 'Valor','Estat'])
    today = stock_data.index[0]
 
    # Alarma 1: Cuando de las ultimas 10 semanas, el indice ha bajado 8
    tabla_calculos, promedios_calc = calculos_table(stock_data)
    var_x_10 = sum(tabla_calculos['Variacio Setmanal'][0:10]>0)
    alarm_dsc = 'De les ultimes 10 setmanes, l\'index ha pujat 8 o mes cops'
    if var_x_10>=8:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor':var_x_10, 'Estat':True}, ignore_index=True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor':var_x_10, 'Estat':False}, ignore_index=True)

    # Alarma 2: Cuando lleva 5 lunes seguidos subiendo
    lun, mar, mie, jue, vie, promedios = tablas_diario(stock_data)
    alarm_dsc = 'Porta 5 dilluns seguits pujant'
    c = 0
    while sum(lun['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)
    
    # Alarma 3: Cuando lleva 5 martes seguidos subiendo
    alarm_dsc = 'Porta 5 dimarts seguits pujant'
    c = 0
    while sum(mar['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 4: Cuando lleva 5 miercoles seguidos subiendo
    alarm_dsc = 'Porta 5 dimecres seguits pujant'
    c = 0
    while sum(mie['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 5: Cuando lleva 5 jueves seguidos subiendo
    alarm_dsc = 'Porta 5 dijous seguits pujant'
    c = 0
    while sum(jue['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 6: Cuando lleva 5 viernes seguidos subiendo
    alarm_dsc = 'Porta 5 divendres seguits pujant'
    c = 0
    while sum(vie['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)
    
    # Alarma 7: Cuando lleva 5 dias seguidos subiendo
    alarm_dsc = 'Porta 5 dies seguits pujant'
    c = 0
    while sum(stock_data['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=5:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 8: Cuando lleva 4 semanas seguidas subiendo
    alarm_dsc = 'Porta 4 setmanes seguides pujant'
    c = 0
    while sum(tabla_calculos['Variacio Setmanal'].head(c+1)>0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 9: Cuando lleva 4 meses seguidos subiendo
    alarm_dsc = 'Porta 4 mesos seguits pujant'
    months = months_data(stock_data)
    c = 0
    while sum(months['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 10: Cuando lleva 4 trimestres seguidos subiendo
    alarm_dsc = 'Porta 4 trimestres seguits pujant'
    trims = trim_data(stock_data)
    c = 0
    while sum(trims['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)

    # Alarma 11: Cuando lleva 4 años seguidos subiendo
    alarm_dsc = 'Porta 4 anys seguits pujant'
    years = years_data(stock_data)
    c = 0
    while sum(years['Diff%'].head(c+1)>0)==c+1:
        c+=1
    if c>=4:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': c, 'Estat': False}, ignore_index = True)
    
    # Alarma 12: Cuando el mes actual ha superado su media a la alza
    alarm_dsc = 'El mes actual ha superat la seva mitja a la alça'
    table_ym = year_month_table(stock_data)
    total_year_ym, alza_baja_ym, medias_ym, promedios_ym, meses_porcentajes_ym = year_month_aux(table_ym)
    if months['Diff%'][0]>promedios_ym.loc['Promedio Subida'][today.month-1]:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': round(months['Diff%'][0],2), 'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': round(months['Diff%'][0],2), 'Estat': False}, ignore_index = True)
    
    # Alarma 13: Cuando la semana actual ha superado el porcentaje de variacion de las ultimas 10 semanas
    alarm_dsc = 'La setmana actual ha superat la mitja de variació de les últimes 10 setmanes a la alça'
    if tabla_calculos['Variacio Setmanal'][0]>tabla_calculos['Porc. Var. 10'][0]:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': round(tabla_calculos['Variacio Setmanal'][0],2),\
                                'Estat': True}, ignore_index = True)
    else:
        alarms = alarms.append({'Alarma Pujada':alarm_dsc, 'Valor': round(tabla_calculos['Variacio Setmanal'][0],2),\
                                'Estat': False}, ignore_index = True)
    return alarms