import os
from tkinter import filedialog
from main import progon_telegi
import PySimpleGUI as sg
import shutil
from os import listdir

def zagr_j_s():
    filepath_list = filedialog.askopenfilenames()
    return filepath_list

def save_word():
    filepath_save = filedialog.askdirectory()
    tmp_list = listdir('TMP_word')
    for files in tmp_list:
        shutil.copy(f'TMP_word/{files}', filepath_save)
        os.remove(f'TMP_word/{files}')


if __name__ == "__main__":
    sg.theme('DarkAmber')

    layout = [  [sg.Text('Все просто, одна кнопка загрузить, другая сохранить')],
                [sg.Button('Загрузить json'), sg.Text('Можно несколько через shift')],
                [sg.Button('Сохранить word'), sg.Button('Выйти')] ]


    window = sg.Window('Стэтхэм в ворд', layout)

    while True:

        event, values = window.read()
        if event == 'Выйти' or event == sg.WIN_CLOSED:
            break

        if event == 'Загрузить json':
            i_num = 1
            list_of_jsons_s = zagr_j_s()
            for element in list_of_jsons_s:
                if '.json' in element:
                    progon_telegi(element, i_num)
                    i_num += 1
                else:
                    print(element)

        if event == 'Сохранить word':
            save_word()

    window.close()

