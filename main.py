import pandas as pd
import PySimpleGUI as sg
import logging
import re
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


def compare_files(path_old, path_new, path_total):
    pd.set_option('display.max_columns', 6)
    df_q = pd.read_excel(path_old, sheet_name='TDSheet', header=7, converters={'Артикул': str}, usecols="B,C")
    df_q = df_q[df_q['Артикул'].notna()]
    # print(df_q.head(20))
    df_a = pd.read_excel(path_new, sheet_name='TDSheet', header=7, converters={'Артикул': str}, usecols="B,C")
    df_a = df_a[df_a['Артикул'].notna()]

    # print(df_a.head(20))

    df_q_a = pd.merge(df_q, df_a, on=['Артикул'], how="outer", indicator=True,
                      suffixes=(" в старом файле", " в новом файле"))
    # print('Merged')
    # print(df_q_a.head(20))
    df_q_a = df_q_a[df_q_a['_merge'] != 'both']
    df_q_a.loc[df_q_a['_merge'] == 'left_only', 'Результат сравнения'] = 'Была только в старом'
    df_q_a.loc[df_q_a['_merge'] == 'right_only', 'Результат сравнения'] = 'Появилась в новом'
    df_q_a.drop('_merge', axis=1, inplace=True)
    # df_q_a.rename({'_merge': 'Результат сравнения', 'Номенклатура_x': 'Номенклатура в старом файле',
    #                'Номенклатура_y': 'Номенклатура в новом файле'}, axis=1, inplace=True)
    # print('Result')
    # print(df_q_a)
    df_q_a.to_excel(path_total, index=False)


def compare_ozon(path_old, path_new, path_total):
    pd.set_option('display.max_columns', 6)
    df_q = pd.read_excel(path_old, sheet_name='TDSheet', header=7, converters={'Артикул': str}, usecols="B,C")
    df_q = df_q[df_q['Артикул'].notna()]
    # print(df_q.head(20))
    df_a = pd.read_excel(path_new, sheet_name='Шаблон для поставщика', header=1, converters={'Артикул*': str},
                         usecols="B,C", skiprows=[2])
    df_a = df_a[df_a['Артикул*'].notna()]
    df_a.rename({'Артикул*': 'Артикул'}, axis=1, inplace=True)

    re_express = re.compile('.*-(.*)-.*')
    df_a['Артикул'] = df_a['Артикул'].str.replace(re_express, r'\1')
    # print(df_a.head(20))

    df_q_a = pd.merge(df_q, df_a, on=['Артикул'], how="inner", indicator=True)
    # print('Merged')
    # print(df_q_a.head(20))


    # df_q_a = df_q_a[df_q_a['_merge'] != 'both']
    # df_q_a.loc[df_q_a['_merge'] == 'left_only', 'Результат сравнения'] = 'Была только в старом'
    # df_q_a.loc[df_q_a['_merge'] == 'right_only', 'Результат сравнения'] = 'Появилась в новом'
    df_q_a.drop('_merge', axis=1, inplace=True)
    # # df_q_a.rename({'_merge': 'Результат сравнения', 'Номенклатура_x': 'Номенклатура в старом файле',
    # #                'Номенклатура_y': 'Номенклатура в новом файле'}, axis=1, inplace=True)
    # print('Result')
    # print(df_q_a)
    df_q_a.to_excel(path_total, index=False)


layout = [

    [sg.Text('Путь к старому файлу остатков Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('Остатки для Семена 20210714.xlsx', size=(64, 1)), sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],
    [sg.Text('Путь к новому файлу остатков/озон Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('2021-11-03 Эротический набор.xlsx', size=(64, 1)), sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],
    [sg.Text('Путь к итоговому файлу Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('Разницы.xlsx', size=(64, 1)),
     sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],

    [sg.Button('Сравнить файлы остатков', key=f'btnRefresh', size=(22, 1)),
     sg.Button('Сравнить остатки с файлом Озона', key=f'btnOzon', size=(28, 1))],

    [sg.Output(size=(112, 12), key='-OUTPUT-')],
]
win = sg.Window('Программа сравнения остатков и Озон', layout, finalize=True)

# ---------
# MAIN LOOP
# ---------
while True:
    event, values = win.read()

    if event == sg.WIN_CLOSED or event == 'Exit':
       break
    elif event == 'btnRefresh':
        try:
           compare_files(values[0], values[1], values[2])
           print("Обработка завершена.")
        except Exception as err:
            print(err)
    elif event == 'btnOzon':
        try:
            compare_ozon(values[0], values[1], values[2])
            print("Обработка завершена.")
        except Exception as err:
            print(err)
    else:
        logger.info(f'This event ({event}) is not yet handled.')
