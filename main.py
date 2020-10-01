from docxtpl import DocxTemplate, InlineImage
import PySimpleGUI as sg
from json import load, dump
from os import system
sg.theme('Dark Blue 3')
# читаем кэш
ImpValues = load(open("UserData", encoding='utf-8'))

"""
Переменные:
1  : НомерРаботы
2  : ТемаРаботы
3  : автор
4  : цельРаботы
5  : постановкаЗадачи
6  : описаниеВарианта
7  : ТеорияПоТеме
8  : структурыКода
9  : интерфейс
10 : описаниеТестирования
11 : рисунокТеста
12 : Освоил : вывод1
13 : Научился: вывод2
14 : исходныйКод
"""
layout = ([sg.Text('Номер Работы:'), sg.Input(ImpValues['0'])],  # 0
          [sg.Text('Тема работы:'), sg.Input(ImpValues['1'])],  # 1
          [sg.Text('Автор:'), sg.Input(ImpValues['2'])],  # 2
          [sg.Text('Цель работы:')],
          [sg.Multiline(ImpValues['3'])],  # 3
          [sg.Text('Постановка задачи:')],
          [sg.Multiline(ImpValues['4'])],  # 4
          [sg.Text('описание Варианта:')],
          [sg.Multiline(ImpValues['5'])],  # 5
          [sg.Text('Теория По Теме:')],
          [sg.Multiline(ImpValues['6'])],  # 6
          [sg.Text('структуры Кода:')],
          [sg.Multiline(ImpValues['7'])],  # 7
          [sg.Text('интерфейс (картинка):'), sg.Text(
              ImpValues['Открыть']), sg.FileBrowse('Открыть')],
          [sg.Text('что тестируем:')],
          [sg.Multiline(ImpValues['8'])],  # 8
          [sg.Text('рисунокТеста:'), sg.Text(
              ImpValues['Открыть0']), sg.FileBrowse('Открыть')],
          [sg.Text('Освоил:'), sg.Input(ImpValues['9'])],  # 9
          [sg.Text('Научился:'), sg.Input(ImpValues['10'])],  # 10
          [sg.Text('исходный Код:'), sg.Text(
              ImpValues['Открыть1']), sg.FileBrowse('Открыть')],
          [sg.Button('Сделать красиво')])
window = sg.Window(
    'АвтоОтчетер (АОД edition) v1 by D.Makarov', layout, icon="ico.ico")
# считываем данные после прочтения
vent, Values = window.read()
# пишем кэш пользователя

for (words, importWords) in zip(Values, ImpValues):
    if len(Values[words]) == 0:
        Values[words] = ImpValues[importWords]

dump(Values, open("UserData", 'w', encoding='utf-8'))

doc = DocxTemplate("example.docx")

# переносим картинки
# обрабатываем полученные данные
# 1

Values['Открыть'] = Values['Открыть'].replace('/', '\\')
InterfaceImage = InlineImage(doc, Values['Открыть'])

# 2

Values['Открыть0'] = Values['Открыть0'].replace('/', '\\')
TestImage = InlineImage(doc, Values['Открыть0'])

# меняем отступ на нужный  для word
Values[7] = Values[7].replace("\n", "<w:br/>")

# копируем наш код из файлика
f = open(Values['Открыть1'].replace('/', '\\'),
         'r', encoding='utf-8')
# меняем запрещенные символы на допустимые для word
Values['Открыть1'] = f.read().replace('<', '&lt;').replace(
    '>', '&gt;').replace("\n", "<w:br/>")
f.close()

# заполняем документ
context = {'НомерРаботы': Values[0],
           'ТемаРаботы': Values[1],
           'автор': Values[2],
           'цельРаботы': Values[3],
           'постановкаЗадачи': Values[4],
           'описаниеВарианта': Values[5],
           'ТеорияПоТеме': Values[6],
           'структурыКода': Values[7],
           'интерфейс': InterfaceImage,
           'описаниеТестирования': Values[8],
           'рисунокТеста': TestImage,
           'вывод1': Values[9],
           'вывод2': Values[10],
           'исходныйКод': Values['Открыть1'],
           }
# делаем отчет
doc.render(context)
doc.save("Отчет - "+Values[1]+".docx")
system('start'+'Отчет - "+Values[1]+".docx')
window.close()
