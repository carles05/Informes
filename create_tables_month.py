import pandas as pd
import numpy as np
import yfinance as yf


def read_data(ind):  # Checked
    """
    Read data using yahoo finance library
    Extracts data from index = ind
    Adds column indicating day of the week
    Adds column for year, month, and day
    """
    index_ticker = yf.Ticker(ind)
    index_data = pd.DataFrame(index_ticker.history(period="max"))
    if ind == '^IBEX':
        index_data = index_data[index_data.index > '1994']
    index_data['Date'] = index_data.index
    index_data.index.rename('Date_index', inplace=True)
    weekday_map = {0: 'Monday', 1: 'Tuesday', 2: 'Wednesday', 3: 'Thursday', 4: 'Friday'}
    # index_data['WeekDay'] = index_data.Date.dt.dayofweek.map(weekday_map) # Add WeekDay column
    index_data['WeekDay'] = index_data.Date.dt.dayofweek
    index_data['Week'] = index_data.Date.dt.week
    index_data['Year'] = index_data.Date.dt.year
    index_data['Month'] = index_data.Date.dt.month
    index_data['Day'] = index_data.Date.dt.day
    index_data = index_data[::-1]
    index_data['Diff%'] = 100 * (index_data.Close / index_data.Close.shift(-1) - 1)
    # index_data.Date = index_data.Date.astype(str)
    index_data.dropna(inplace=True)
    index_data.drop(columns=['Dividends', 'Stock Splits'], inplace=True)
    return index_data


def read_data_csv(index_data):  # Checked
    """
    Read data using yahoo finance library
    Extracts data from index = ind
    Adds column indicating day of the week
    Adds column for year, month, and day
    """
    index_data['Date'] = index_data.index
    weekday_map = {0: 'Monday', 1: 'Tuesday', 2: 'Wednesday', 3: 'Thursday', 4: 'Friday'}
    # index_data['WeekDay'] = index_data.Date.dt.dayofweek.map(weekday_map) # Add WeekDay column
    index_data['WeekDay'] = index_data.Date.dt.dayofweek
    index_data['Week'] = index_data.Date.dt.week
    index_data['Year'] = index_data.Date.dt.year
    index_data['Month'] = index_data.Date.dt.month
    index_data['Day'] = index_data.Date.dt.day
    index_data = index_data[::-1]
    index_data['Diff%'] = 100 * (index_data.Close / index_data.Close.shift(-1) - 1)
    # index_data.Date = index_data.Date.astype(str)
    index_data.dropna(inplace=True)
    index_data.drop(columns=['Date'], inplace=True)
    return index_data


def months_data(data):
    """
    Takes stock data
    Returns table with all the months represented
    by the last day of the month
    """
    data_f = data.copy()
    data_f['LDM'] = data_f.Month != data_f.Month.shift(1)
    data_f = data_f[data_f['LDM'] == True]
    data_f['Diff%'] = 100 * (data_f.Close / data_f.Close.shift(-1) - 1)
    return data_f


def year_month_table(data):
    """
    Takes stock data
    Returns a table with index = Year and columns = Month
    with the percentage changes respect to previous month
    """
    data_f = data.copy()
    data_f['LDM'] = data_f.Month != data_f.Month.shift(1)
    data_f = data_f[data_f['LDM'] == True]
    data_f.drop(columns=['LDM'], inplace=True)
    data_f['Diff%'] = 100 * (data_f.Close / data_f.Close.shift(-1) - 1)
    data_f = pd.pivot_table(data_f, values='Diff%', index='Year', columns='Month'). \
        sort_values(by='Year', ascending=False)
    months_map = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
                  7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}
    data_f.rename(columns=months_map, inplace=True)
    return data_f


def year_month_aux(data):
    # Total año de tabla principal
    total_year = data.sum(axis=1)
    # Meses alza / Meses baja
    alza = data[data > 0].count(axis=1)
    baja = data[data < 0].count(axis=1)
    alza_baja = pd.DataFrame({'Meses Alza': alza, 'Meses Baja': baja})
    # Tabla medias de cada mes
    table_ym_1 = data.copy()
    table_ym_1 = table_ym_1.iloc[1:]
    media_t = table_ym_1.mean(axis=0, skipna=True)
    media_10 = table_ym_1.head(10).mean(axis=0, skipna=True)
    media_11 = table_ym_1.head(11).mean(axis=0, skipna=True)
    l = [np.abs(media_10 * 2), media_t, media_11, media_10]
    medias = pd.concat(l, axis=1).T
    medias = medias.rename(index={0: 'Objetivo Actual', 1: 'Media Total', 2: 'Media ultimos 11', 3: 'Media ultimos 10'})
    medias['Año'] = medias.sum(axis=1)
    # tabla promedios
    promedios = pd.DataFrame()
    promedios['%Sube'] = data[data > 0].count(axis=0) / data.count(axis=0) * 100
    promedios['%Baja'] = data[data < 0].count(axis=0) / data.count(axis=0) * 100
    promedios['Promedio Subida'] = data[data > 0].mean(axis=0)
    promedios['Promedio Bajada'] = data[data < 0].mean(axis=0)
    promedios = promedios.T
    promedios['Año'] = [total_year[total_year > 0].count() / total_year.count() * 100,
                        total_year[total_year < 0].count() / total_year.count() * 100,
                        total_year[total_year > 0].mean(),
                        total_year[total_year < 0].mean()]
    # tabla alza baja porcentajes
    meses_porcentajes = pd.DataFrame({'Frac': ['0/12', '1/11', '2/10', '3/9', '4/8', '5/7', '6/6', '7/6',
                                               '8/4', '9/3', '10/2', '11/1', '12/0'],
                                      'Count': [0] * 13})
    for i in range(alza_baja.shape[0]):
        if alza_baja.iloc[i, 0] + alza_baja.iloc[i, 1] == 12:
            meses_porcentajes.iloc[alza_baja.iloc[i, 0], 1] += 1
    meses_porcentajes['Perc'] = meses_porcentajes['Count'] / meses_porcentajes['Count'].sum() * 100
    meses_porcentajes.drop(columns='Count', inplace=True)
    return total_year, alza_baja, medias, promedios, meses_porcentajes
