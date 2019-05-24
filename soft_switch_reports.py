#%%
import openpyxl as xl
import numpy as np
import os

dir = os.path.abspath(os.curdir)
inDir = dir + '\\InFiles\\test_data.xlsx'
outDir = dir + '\\OutFiles\\only_clear_data.xlsx'
print (inDir)


wb = xl.load_workbook(filename=inDir)

ws = wb.active
data = np.array([[cell.value for cell in row] for row in ws.iter_rows()])
x_str = 'Существенное'


row_index = 0
row_index_list = []
for row in data:
    for cell in row:
        if cell == 'Существенное Значение':
            # print(cell)
            row_index_list.append(row_index)
    row_index = row_index + 1
row_index_list.reverse()

for x in row_index_list:
    data = np.delete(data, x, axis = 0)


res_wb = xl.Workbook()
res_ws = res_wb.active

for row in data:
    res_ws.append(row.tolist())

res_wb.save(outDir)
print ('Script completed')
