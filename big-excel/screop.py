import pandas as pd
from scfunc import *

output_cols = [
  'LABEL(-)',
  'REGIÓN',
  'ZONA',
  'AGENCIA',
  'COMITÉ',
  'ANALISTA',
  'Nombre',
  'Categoría',
  'Tipo',
  'Días Trabajados',
  'Fecha Ingreso',
  'Saldos al día Clientes',
  'Saldos al día Monto',
  'Saldos al día Saldo Promedio',
  'Saldos al día Mora SBS',
  'Variación Anual Cartera Bruta Clientes',
  'Variación Anual Cartera Bruta Monto',
  'Variación Mensual Cartera Bruta Clientes',
  'Variación Mensual Cartera Bruta Monto',
  'Variación Día Cartera Bruta Clientes',
  'Variación Día Cartera Bruta Monto',
  'Nuevos Nº',
  'Nuevos Monto',
  'Recurrentes Nº',
  'Recurrentes Monto',
  'Total Nº',
  'Total Monto',
  'Micro Monto',
  'Micro %',
  'TEA %',
  'Cartera al día (0 días) Saldo',
  'Cartera al día (0 días) %',
  'Mora de 1-8 Saldo',
  'Mora de 1-8 %',
  'Mora 9-30 Saldo',
  'Mora 9-30 %',
  'Mora > 30 Saldo',
  'Mora > 30 %',
  'Mora > 8 Saldo',
  'Mora > 8 %',
  'Mora > 8 + Castigos + Venta Saldo',
  'Mora > 8 + Castigos + Venta %',
  'Variación Mora > 8 + Castigos + Venta Saldo',
  'Variación Mora > 8 + Castigos + Venta %',
  'Variación Mora > 30 + Castigos + Venta Saldo',
  'Variación Mora > 30 + Castigos + Venta %',
  'Tramo de 1 a 30 Base a contener (inicio de mes)',
  'Tramo de 1 a 30 Cartera Contenida',
  'Tramo de 1 a 30 Cartera Deteriorada',
  'Tramo de 1 a 30 % Monto Contenido',
  'Tramo de 31 a 60 Base a contener (inicio de mes)',
  'Tramo de 31 a 60 Cartera Contenida y Liberada',
  'Tramo de 31 a 60 Cartera Deteriorada',
  'Tramo de 31 a 60 % Monto Contenido y Liberado',
  'Cartera Bruta Meta Mensual',
  'Cartera Bruta Logro Mensual (%)',
  'Clientes Totales Meta Mensual',
  'Clientes Totales Logro Mensual (%)',
  'Mora Contable Meta',
  'Mora Contable Logro (%)',
  'Retención de Clientes Nº Base',
  'Retención de Clientes Nº Retenidos',
  'Retención de Clientes % Retención',
  'Retencion de Clientes de Alto Valor Nº Base',
  'Retencion de Clientes de Alto Valor Nº Retenidos',
  'Retencion de Clientes de Alto Valor % Retención',
  'Nº Pagos Programados del Mes',
  'Nº Pagos Programados al Día',
  'Nº Pagos Ejecutados a Hoy',
  '% Pago a Hoy',
  'Crecimiento Saldo',
  'Crecimiento Clientes',
  'Retención',
  'Contención de 31 a 60',
  'Cartera > a 8',
  'Productividad',
  'Puntaje Total',
  'Calificación'
]

mydir = 'shared/'
filename = 'REP_R327_OPERATIVO_GENERAL_RRHH_20240828.xlsx'

big_table = pd.read_excel(mydir + filename, 
                          header=[4,5,6,7,8,9], index_col=0)

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)

caja_arequipa_row = big_table.columns.get_level_values(-1)
big_table.columns = big_table.columns.droplevel(-1)
print(caja_arequipa_row)
print('*'*100)

column_ranges = [(1, 6), (6, 16), (16, 25), (25, 41), (41, 49), (49, 61),
                 (61, 65), (65, 73)]
group_names = ['Datos Analista', 'Cartera Bruta', 'Desembolso Acumulado Mensual',
               'Calidad de Cartera', 'Gestion por Tramos de Cartera',
               'Retencion de Clientes', 'Seguimiento de Pagos', 'Puntaje']

big_table_d = split_df_by_column_ranges(big_table, column_ranges, group_names)

big_table_d['Datos Analista'] = big_table_d['Datos Analista'].droplevel(level=[0,1,2,3], axis=1)
big_table_d['Cartera Bruta'] = big_table_d['Cartera Bruta'].droplevel(level=[0,1,3], axis=1)
big_table_d['Desembolso Acumulado Mensual'] = big_table_d['Desembolso Acumulado Mensual'].droplevel(level=[0,1,2], axis=1)
big_table_d['Calidad de Cartera'] = big_table_d['Calidad de Cartera'].droplevel(level=[0,1,3], axis=1)
big_table_d['Gestion por Tramos de Cartera'] = big_table_d['Gestion por Tramos de Cartera'].droplevel(level=[0,1], axis=1)#
big_table_d['Retencion de Clientes'] = big_table_d['Retencion de Clientes'].droplevel(level=[0,1,3], axis=1)
big_table_d['Seguimiento de Pagos'] = big_table_d['Seguimiento de Pagos'].droplevel(level=[0,1,2], axis=1)#
big_table_d['Puntaje'] = big_table_d['Puntaje'].droplevel(level=[0,1,2], axis=1)#

#print(big_table_d['Datos Analista'][' - TOTAL CAJA']['REGION ANDINA'].head(15))

for group in group_names:
  big_table_d[group] = merge_multicolumns(big_table_d[group])
  big_table_d[group] = create_hierarchical_rows(big_table_d[group])
  print(big_table_d[group].head(30))
  print('*'*100)

n = 6
pre_colx = list(big_table_d.values())[0].iloc[:n]

output = pd.concat([pre_colx] + list(big_table_d.values()), ignore_index=True)

output = flatten_rows(output, n)

output.columns = output_cols

output = output.dropna(subset=['Cartera Bruta Meta Mensual'])

output.to_excel(mydir + '(mod)' + filename, index=None)