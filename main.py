import pandas as pd
from xlsxwriter.workbook import Workbook

import PySimpleGUI as sg

import logging


logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


def export_to_excel(path):
    try:
        workbook = Workbook(path)
        worksheet = workbook.add_worksheet()
        conn = sqlite3.connect('sqlite.db')
        c = conn.cursor()
        mysel = c.execute("select * from Journal")
        for i, row in enumerate(mysel):
            for j, value in enumerate(row):
                worksheet.write(i, j, value)
        workbook.close()
        print("Данные успешно экспортированы в файл ", path)
    except Exception as err:
        print(err)


layout = [

    [sg.Text('Путь к старому файлу Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('Остатки для Семена 20210714.xlsx', size=(64, 1)), sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],
    [sg.Text('Путь к новому файлу Excel:', size=(35, 1), auto_size_text=False, justification='left'),
     sg.InputText('Остатки для Семена 20210723.xlsx', size=(64, 1)),
     sg.FileBrowse(file_types=(("Excel files", "*.xlsx"),))],

    [sg.Button('Сравнить файлы', key=f'btnRefresh', size=(22, 1)),
     sg.Button('Экспортировать в Excel', key=f'btnExport', size=(22, 1))],

    [sg.Output(size=(112, 6), key='-OUTPUT-')],
]
win = sg.Window('Программа сравнения остатков', layout, finalize=True)

# ---------
# MAIN LOOP
# ---------
while True:
    event, values = win.read()

    if event == sg.WIN_CLOSED or event == 'Exit':
       break
    elif event == 'btnRefresh':
        try:
           print("Обработка завершена.")
        except Exception as err:
            print(err)
    elif event == 'btnExport':
        export_to_excel(values[3])
    else:
        logger.info(f'This event ({event}) is not yet handled.')
