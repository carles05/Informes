# Import libraries and functions to create tables
import numpy as np
import yfinance as yf
import pandas as pd
from datetime import date
from create_tables_month import read_data, year_month_table, year_month_aux
from create_tables_trim import year_trim_table, year_periods, year_trim_aux
from create_table_calculos import calculos_table
from create_tables_notas import notas_tables
from create_resumen_diario import tablas_diario
from create_alarms import alarms_baja, alarms_alza
from cuadro_control import CreateDashboard

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
    worksheet.set_column('C:C', 20, centered)
    worksheet.set_column('D:D', 20, centered)
    worksheet.set_column('E:E', 20, centered)
    worksheet.set_column('F:F', 20, centered)
    worksheet.set_column('G:G', 20, centered)


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
    notas_tabs[years[0]].to_excel(writer, sheet_name = 'Notas', startrow = 1, startcol = 1)
    notas_tabs[years[1]].to_excel(writer, sheet_name = 'Notas', startrow = 16, startcol = 1)
    notas_tabs[years[2]].to_excel(writer, sheet_name = 'Notas', startrow = 31, startcol = 1)
    worksheet = writer.sheets['Notas']
    worksheet.write(1, 1, str(years[0]))
    worksheet.write(16, 1, str(years[1]))
    worksheet.write(31, 1, str(years[2]))
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
    writer.save()

num = 1
while num!=0:
    num = int(input('Quin informe vols generar?\n1.IBEX\n2.NASDAQ\n3.DOWJONES\n4.SP500\n5.E50\n0.EXIT\n--> '))
    # Read tables
    if num!=0:
        inds = ["^IBEX", "^IXIC", "^DJI", "^GSPC", "^STOXX50E"]
        ind = inds[num-1]
        stock_dict = {"^IBEX":"IBEX", "^IXIC":"NASDAQ",\
                    "^DJI":"DOWJONES", "^GSPC": "SP500", "^STOXX50E": "E50"}
        generarExcel(ind)
