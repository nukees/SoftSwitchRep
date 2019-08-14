# -*- coding: utf-8 -*-

# Базовые библиотеки
import os
import sys
import re
import math
from datetime import datetime

# Скачанные библиотеки (pandas, xlrd, numpy, openpyxl)
import pandas as pd
import openpyxl as oxl
from openpyxl.utils.dataframe import dataframe_to_rows

def clearAndSortTable (file_with_table):
    df = pd.read_excel('IN\\'+ file_with_table, sheet_name = 'table1', skiprows = 6)
    df = df[(df != 'Significant Value').all(axis=1)]
    df = df.sort_values(by = ['Trunk Group (ID)','by Hour']).reset_index(drop = True)
    columns = df.columns.tolist()
    x = columns[0]
    del columns[0]
    columns.insert(2, x)
    df = df[columns]
    return df

def findNormalizedCoeff (number_all_chanels):
    list_coeff_values = [
                        0.01, 0.07, 0.15, 0.22, 0.27, 0.32, 0.35, 0.39, 0.42, 0.44, 
                        0.46, 0.49, 0.5, 0.52, 0.54, 0.55, 0.56, 0.57, 0.59, 0.6,
                        0.61, 0.61, 0.62, 0.63, 0.64, 0.65, 0.65, 0.66, 0.67, 0.67,
                        0.68, 0.68, 0.69, 0.69, 0.7, 0.7, 0.71, 0.71, 0.71, 0.72, 
                        0.72, 0.73, 0.73, 0.73, 0.74, 0.74, 0.74, 0.74, 0.75, 0.75, 
                        0.75, 0.76, 0.76, 0.76, 0.76, 0.77, 0.77, 0.77, 0.77, 0.78, 
                        0.78, 0.78, 0.78, 0.78, 0.78, 0.79, 0.79, 0.79, 0.79, 0.79, 
                        0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8
                        ]
    if (number_all_chanels > 77):
        x = 0.81
    else:
        x = list_coeff_values[number_all_chanels - 1]
    return x

def findNeedChanels(coeff_norm, total_load):
    list_k = [
            0.01,0.15,0.46,0.86,1.35,1.86,2.48,3.1,3.74,4.42,
            5.11,5.82,6.54,7.28,8.03,8.79,9.55,10.34,11.12,
            11.91,12.71,13.51,14.33,15.15,15.97,16.79,17.62,
            18.45,19.3,20.14,20.98,21.83,22.68,23.53,24.39,
            25.26,26.12,26.98,27.85,28.72,29.59,30.46,31.34,
            32.22,33.1,33.98,34.87,35.75,36.63,37.52,38.41,
            39.3,40.19,41.09,41.99,42.88,43.78,44.68,45.58,
            46.48,47.38,48.28,49.19,50.09,51.01,51.92,52.82,
            53.73,54.64,55.55,56.46,57.38,58.29,59.2,60.12,
            61.03,61.95,62.88,63.79,64.71,65.61,66.42,67.23,
            8.04,68.85,69.66,70.47,71.28,72.09,72.9,73.71,
            74.52,75.33,76.14,76.95,77.76,78.57,79.38,80.19,
            81,81.81,82.62,83.43,84.04,85.05,85.86,86.67,87.48,
            88.29,89.1,89.91,90.72,91.53,92.34,93.15,93.96,
            94.77,95.58,96.36,97.2,98.01,98.82,99.63,100.44,
            101.25,102.06,102.87,103.68,104.49,105.3,106.11,
            106.92,107.73,108.54,109.35,110.16,110.97,111.38,
            112.59,113.4,114.21,115.02,115.83,116.64,117.45,
            118.26,119.07,119.88,120.69,121.5,122.31,123.12,
            123.93,124.74,125.55,126.36,127.17,127.98,128.79,
            129.6,130.41,131.22,132.04,132.84,133.65,134.46,
            135.27,136.08,136.89,137.7,138.51,139.32,140.13,
            140.94,141.75,142.56,143.37,144.18,144.99,145.8,
            146.61,147.42,148.23,149.04,149.85,150.66,151.49,
            152.28,153.09,153.9,154.71,155.52,156.33,157.14,
            157.95,158.76,159.57,160.38,161.19,162
        ]
    if (coeff_norm == 0.81):
            need_chanels = math.ceil(total_load/coeff_norm)
    else:
        if (total_load == 0):
                need_chanels = 0
        else:
            for step in range(len(list_k)):
                if  (total_load < list_k[step]):
                    need_chanels = step + 1
                    break
    return need_chanels

def createSheet(wb, sheet_name, df):
    wb.create_sheet(sheet_name)
    ws = wb[sheet_name]
    rows, _ = df.shape
    if (rows > 0):
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
    else:
        ws.append(['Данные отсутствуют'])

def createReport (file_out, file_in):
    num = file_out[:3]

    df_out = clearAndSortTable(file_out)
    df_in = clearAndSortTable(file_in)
  
    res_df_6A = pd.DataFrame()
    res_df_6B = pd.DataFrame()
    temp_df = pd.DataFrame()

    temp_df['11'] = df_out['All Channels'].apply(findNormalizedCoeff)
    temp_df['Total'] = df_out['Traffic Volume - PeakHour'] + df_in['Traffic Volume - PeakHour']
    temp_df['9'] = temp_df[['11','Total']].apply(lambda x: findNeedChanels(x['11'], x['Total']), axis = 1)

    res_df_6A['1'] = df_out['Trunk Group (ID)']
    res_df_6A['2'] = df_out['Trunk Group (Name)']
    res_df_6A['3'] = df_out['by Hour']
    res_df_6A['4'] = df_out['Traffic Volume - PeakHour'] # + df_in['Traffic Volume - PeakHour']
    res_df_6A['5'] = df_out['PeakHour Overflow to All Ratio']
    res_df_6A['6'] = df_out['PeakHour Number Of Seizures']
    res_df_6A['7'] = df_out['All Channels']
    res_df_6A['8'] = df_out['Active Channels']
    res_df_6A['9'] = res_df_6A['7'] - temp_df['9']
    res_df_6A['10'] = df_out['PeakHour Channels Utilization']
    res_df_6A['11'] = temp_df['11']
    res_df_6A['12'] = '-'
    res_df_6A['13'] = '-'
    res_df_6A['14'] = df_out['PeakHour Average Seizure Length']

    res_df_6B['1'] = df_in['Trunk Group (ID)']
    res_df_6B['2'] = df_in['Trunk Group (Name)']
    res_df_6B['3'] = df_in['by Hour']
    res_df_6A['4'] = df_in['Traffic Volume - PeakHour'] # + df_in['Traffic Volume - PeakHour']
    res_df_6B['5'] = df_in['PeakHour Number Of Seizures']
    res_df_6B['6'] = df_in['All Channels']
    res_df_6B['7'] = df_in['Active Channels']
    res_df_6B['8'] = res_df_6B['6'] - temp_df['9']
    res_df_6B['9'] = df_in['PeakHour Channels Utilization']
    res_df_6B['10'] = temp_df['11']
    res_df_6B['11'] = df_in['PeakHour Average Seizure Length']
     
    wb = oxl.Workbook()
    report_date = datetime.now()

    createSheet(wb, 'Форма 6А', res_df_6A)
    createSheet(wb, 'Форма 6B', res_df_6B)

    del wb['Sheet']
    wb.save('OUT\\{}_report_{}.xlsx'.format(num, report_date.strftime('%Y-%m-%d')))
            

if __name__ == "__main__":
    all_in_files = os.listdir('IN')
    need_in_files = []
    softswitchs_numbers = ['712', '715', '714', '716', '729', '726']
    pairs_switches = []

    # Фильтрация файлов с помощью регулярных выражений
    # https://proglib.io/p/learn-regex/

    filtered_files = [f for f in all_in_files if (re.search(r'\d{3}_[A-Za-z]{,3}_\d{4}-\d{2}-\d{2}.xlsx', f))]
    # filtered_in_files = [f for f in filtered_files if (re.search(r'(?:IN|in)|(?:OUT|out)',f))]

    for switch_num in softswitchs_numbers:
        pair = []
        for f in filtered_files:
            if (switch_num == f[:3]):
                pair.append(f)
        if (len(pair) == 2):
            pairs_switches.append(pair)
            
    for p in pairs_switches:
        file_in = p[0]
        file_out = p[1]
        createReport(file_out, file_in)
    # createReport(file_out, file_in)
    print('--Обработано--')