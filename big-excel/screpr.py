import pandas as pd
from scfunc import *

output_cols = [
  'LABEL(-)',
  'REGIÓN',
  'ZONA',
  'AGENCIA',
  'COMITÉ',
  'ANALISTA',
  'Total Operaciones Productivas KPI'
]

mydir = 'shared/'
filename = 'REP_R017_PRODUCTIVIDAD_20240828.xlsx'

big_table = pd.read_excel(mydir + filename, 
                          header=[5,6,7,8,9], index_col=0)

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)

caja_arequipa_row = big_table.columns.get_level_values(-1)
big_table.columns = big_table.columns.droplevel(-1)
print(caja_arequipa_row)
print('*'*100)

print(big_table.head(30))
print('*'*100)

column_ranges = [(6,7)]
group_names = ['KPI de Productividad']

big_table_d = split_df_by_column_ranges(big_table, column_ranges, group_names)
print(big_table_d['KPI de Productividad'].head(15))
print('*'*100)
big_table_d['KPI de Productividad'] = big_table_d['KPI de Productividad'].droplevel(level=[1,3], axis=1)

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

output.to_excel(mydir + '(mod)' + filename, index=None)