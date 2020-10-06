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
        pass
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

tab_html = ()

layout = [[sg.TabGroup([[sg.Tab('АОД', tab_aod),
                         sg.Tab('HTML', tab_html)]])]]
window = sg.Window(
    'АвтоОтчетер  v1.1 by D.Makarov', layout, icon="ico.ico")
# read input
while True:
    event, Values = window.read()
    event = int(event.replace(".Сделать красиво"))
    print("event "+str(type(event))+": "+str(event))
    doc = DocxTemplate("examples/аод example.docx")
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
            f = open(Values[word], 'r', encoding='utf-8')
            # change forbidden characters to valid for word
            Values[word] = [f.read().replace('<', '&lt;').replace(
                '>', '&gt;').replace("\n", "<w:br/>"), Values[word]]
            f.close()

    # fill out the document
    context = makecontext(event, Values)

    # make a report
    doc.render(context)
    doc.save("Отчет - "+context['темаРаботы']+".docx")
    # replace img to way
    for word in Values:
        if(type(Values[word]) == list):
            Values[word] = Values[word][1]
    # save cache
    dump(Values, open("UserData.json", 'w', encoding='utf-8'))
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
window.close()
