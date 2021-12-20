import pandas as pd
import PySimpleGUI as sg
import logging
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


def compare_files(path_old, path_new, path_total):
    pd.set_option('display.max_columns', 6)
    df_q = pd.read_excel(path_old, header=0, converters={'Номер отправления': str}, usecols="A,B,G,H,I")
    df_q = df_q[df_q['Номер отправления'].notna()]
    df_a = pd.read_excel(path_new, header=0, converters={'Номер отправления': str}, usecols="A,B,G,H,I")
    df_a = df_a[df_a['Номер отправления'].notna()]

    df_q_a = pd.merge(df_q, df_a, on=['Номер отправления'], how="outer", indicator=True,
                      suffixes=(" в старом файле", " в новом файле"))
    df_q_a = df_q_a[df_q_a['_merge'] != 'both']
    df_q_a.loc[df_q_a['_merge'] == 'left_only', 'Результат сравнения'] = 'Было только в старом'
    df_q_a.loc[df_q_a['_merge'] == 'right_only', 'Результат сравнения'] = 'Появилось в новом'
    df_q_a.drop('_merge', axis=1, inplace=True)
    df_q_a.to_excel(path_total, index=False)


layout = [

    [sg.Text('Путь к старому файлу отправлений Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('Сем1.xlsx', size=(64, 1)), sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],
    [sg.Text('Путь к новому файлу отправлений Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('Сема2.xlsx', size=(64, 1)), sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],
    [sg.Text('Путь к итоговому файлу Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('Разницы.xlsx', size=(64, 1)),
     sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],

    [sg.Button('Сравнить файлы отправлений', key=f'btnRefresh', size=(26, 1))],

    [sg.Output(size=(112, 12), key='-OUTPUT-')],
]
win = sg.Window('Программа сравнения отправлений Озон', layout, finalize=True)

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
    else:
        logger.info(f'This event ({event}) is not yet handled.')
