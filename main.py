import pandas as pd
from openpyxl import Workbook, load_workbook


def calc_percentage(new_cell, old_cell):
    percentage = ((new_cell - old_cell)/old_cell)*100

def create_percentage_table():
    pass


df = pd.read_excel('cmdb_ci_db_instance.xlsx')
num_filas = df.shape[0]
num_columnas = df.shape[1]

wb = load_workbook('Libro1.xlsx')
ws = wb.active
ws['B3'] = num_filas
wb.save('Libro2.xlsx')
df = pd.read_excel('Libro2.xlsx')
num_filas_res = df.shape[0]
num_columnas_res = df.shape[1]
#for i in range(num_filas_res):

print("Excel de entrada:", num_filas, num_columnas)
print("Excel de salida:", num_filas_res,num_columnas_res)

