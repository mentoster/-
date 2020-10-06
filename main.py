from docxtpl import DocxTemplate, InlineImage
import PySimpleGUI as sg
from json import load, dump
# theme
sg.theme('BlueMono')
# read the cache
ImpValues = load(open("UserData.json", encoding='utf-8'))


def makecontext(tab, Values):
    context = {}
    if tab == 1:
        context = {'НомерРаботы': Values[0],
                   'ТемаРаботы': Values[1],
                   'автор': Values[2],
                   'цельРаботы': Values[3],
                   'постановкаЗадачи': Values[4],
                   'описаниеВарианта': Values[5],
                   'ТеорияПоТеме': Values[6],
                   'структурыКода': Values[7],
                   'интерфейс': Values['Open'][0],
                   'описаниеТестирования': Values[8],
                   'рисунокТеста': Values['Open0'][0],
                   'вывод1': Values[9],
                   'вывод2': Values[10],
                   'исходныйКод': Values['Open1'][0],
                   }
    elif tab == 2:
        context = {'НомерРаботы': Values[11],
                   'ТемаРаботы': Values[12],
                   'автор': Values[2],
                   'Название': Values[13],
                   'задание': Values[14],
                   'описание': Values[15],
                   'код': Values['Open2'][0],
                   'css': Values['Open3'][0],
                   'рисунок': Values['Open4'][0],
                   'вывод': Values[16],
                   }
    return context


tab_aod = ([sg.Text('Номер Работы:'), sg.Input(ImpValues['0'])],  # 0
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
               ImpValues['Open']), sg.FileBrowse('Open')],
           [sg.Text('что тестируем:')],
           [sg.Multiline(ImpValues['8'])],  # 8
           [sg.Text('рисунокТеста:'), sg.Text(
               ImpValues['Open0']), sg.FileBrowse('Open')],
           [sg.Text('Освоил:'), sg.Input(ImpValues['9'])],  # 9
           [sg.Text('Научился:'), sg.Input(ImpValues['10'])],  # 10
           [sg.Text('исходный Код:'), sg.Text(
               ImpValues['Open1']), sg.FileBrowse('Open')],
           [sg.Button('1.Сделать красиво')])

tab_html = ([sg.Text('Номер Работы:'), sg.Input(ImpValues['11'])],
            [sg.Text('Тема работы:'), sg.Input(ImpValues['12'])],
            [sg.Text('автор'), sg.Input(ImpValues['2'])],
            [sg.Text('Название практического задания:'),
             sg.Input(ImpValues['13'])],
            [sg.Text('описание задания:')],
            [sg.Multiline(ImpValues['15'])],
            [sg.Text('код'), sg.Text(
                ImpValues['Open']), sg.FileBrowse('Open2')],
            [sg.Text('css код'), sg.Text(
                ImpValues['Open']), sg.FileBrowse('Open3')],
            [sg.Text('рисунок сайта'), sg.Text(
                ImpValues['Open']), sg.FileBrowse('Open4')],
            [sg.Text('вывод :'), sg.Input(ImpValues['16'])],
            [sg.Button('2.Сделать красиво')])

layout = [[sg.TabGroup([[sg.Tab('АОД', tab_aod),
                         sg.Tab('HTML', tab_html)]])]]
window = sg.Window(
    'АвтоОтчетер  v1.1 by D.Makarov', layout, icon="ico.ico")
# read input
while True:
    event, Values = window.read()
    if type(event) == str:
        event = int(event.replace(".Сделать красиво", ""))
    else:
        break
    doc = DocxTemplate("examples/"+str(event)+".docx")

    for (word, importWords) in zip(Values, ImpValues):
        # load cache
        if len(Values[word]) == 0:
            Values[word] = ImpValues[importWords]
        # open imagine
        Values[word] = Values[word].replace('/', '\\')
        if (".png" or ".jpg") in Values[word]:
            Values[word] = [InlineImage(doc, Values[word]), Values[word]]
        # open files
        elif "\\" in Values[word]:
            try:
                f = open(Values[word], 'r', encoding='utf-8')
                # change forbidden characters to valid for word
                Values[word] = [f.read().replace('<', '&lt;').replace(
                    '>', '&gt;').replace("\n", "<w:br/>"), Values[word]]
                f.close()
            except exp as identifier:
                f = open(Values[word], 'r', encoding='utf-8')
                Values[word] = [f.read().replace('<', '&lt;').replace(
                    '>', '&gt;').replace("\n", "<w:br/>"), Values[word]]
                f.close()

    # fill out the document
    context = makecontext(event, Values)
    # make a report
    doc.render(context)
    doc.save("Отчет - "+context['ТемаРаботы']+".docx")
    # replace img to way
    for word in Values:
        if(type(Values[word]) == list):
            Values[word] = Values[word][1]
    # save cache
    # dump(Values, open("UserData.json", 'w', encoding='utf-8'))
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
window.close()
