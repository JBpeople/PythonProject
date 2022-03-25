# -*- coding: utf-8 -*-
# @Time : 2021-12-6 9:10
# @Author : Administrator
# @Email : lcmmljs@foxmail.com
# @File : main.py
# @Project : Project

# -*- coding: utf-8 -*-
# @Time : 2021-12-6 9:10
# @Author : Administrator
# @Email : lcmmljs@foxmail.com
# @File : main.py
# @Project : Project

import glob
import PySimpleGUI as sg
from PyPDF2 import PdfFileReader, PdfFileWriter

# 创建GUI布局
fra1 = sg.Frame('打开文件夹', [[sg.InputText(key='-IN-'), sg.FolderBrowse('文件位置')], [sg.Button('确认')]])
fra2 = sg.Frame('文件列表', [[sg.Listbox(values=[''], s=(54, 30), key='-LIST-')],
                         [sg.Button('上移'), sg.Button('下移'), sg.Button('移除'), sg.Button('合并')]])

col = sg.Column([[fra1], [fra2]], pad=(0, 0))

layout = [[col]]
window = sg.Window('合并PDF', layout, keep_on_top=True)

# 创建程序主循环
while True:
    event, value = window.read()
    if event is None:
        break
    elif event == '确认':
        if value['-IN-'] != '':
            path = window['-IN-'].get() + '/*.pdf'
            file = glob.glob(path)
            if file:
                filefloder = file[0].split('\\')[0] + '/'
                filename = [i.split('\\')[-1] for i in file]
                window['-LIST-'].update(filename)
            else:
                sg.popup_ok('文件夹内未识别到PDF文件~', title='错误', keep_on_top=True)
        else:
            sg.popup_ok('请选择文件夹~', title='错误', keep_on_top=True)
    elif event == '上移':
        try:
            choice = value['-LIST-']
            index = filename.index(choice[0])
            filename[index], filename[index-1] = filename[index-1], filename[index]
            window['-LIST-'].update(filename)
        except NameError:
            pass
    elif event == '下移':
        try:
            choice = value['-LIST-']
            index = filename.index(choice[0])
            filename[index], filename[index+1] = filename[index+1], filename[index]
            window['-LIST-'].update(filename)
        except NameError:
            pass
    elif event == '移除':
        try:
            choice = value['-LIST-'][0]
            filename.remove(choice)
            window['-LIST-'].update(filename)
        except NameError:
            pass
        except IndexError:
            pass
    elif event == '合并':
        try:
            if filename:
                file = [filefloder + i for i in filename]
                pdfFileWriter = PdfFileWriter()
                for pdf in file:
                    pdfReader = PdfFileReader(open(pdf, 'rb'))
                    numPages = pdfReader.getNumPages()
                    for i in range(0, numPages):
                        pageObj = pdfReader.getPage(i)
                        pdfFileWriter.addPage(pageObj)
                folder = sg.popup_get_folder('请选择文件保存位置：', title='提示', keep_on_top=True)
                pdfFileWriter.write(open(folder + '//合并文件.pdf', 'wb'))
                sg.popup_ok('文件已合并完成~', keep_on_top=True)
        except NameError:
            pass
window.close()
