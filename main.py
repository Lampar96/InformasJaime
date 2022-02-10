import pandas as pd
from openpyxl import Workbook, load_workbook

df = pd.read_excel('cmdb_ci_db_instance.xlsx', engine='openpyxl')
num_filas = df.shape[0]

wb = load_workbook('Libro1.xlsx')
ws = wb.active
ws['B3'] = 0
wb.save('Libro2.xlsx')


print(df.shape[0])
