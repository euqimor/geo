from os import getcwd
from appJar import gui
from coords import *

app = gui('Гео', '400x260')
app.setPadding([2,0])
app.setInPadding([0,0])
app.setGuiPadding(4,14)
app.setStretch('column')
app.setFont(12)
app.setBg('lightBlue')

def open_file(button):
    a = app.openBox()
    app.setEntry('addr_path', a)


def save_file(button):
    a = app.saveBox()
    app.setEntry('save_as', a)


def press(button):
    p1 = app.getEntry('addr_path')
    p2 = app.getEntry('save_as')
    with open(p1) as f:
        print(f.readline())
    with open(p2) as f:
        print(f.readline())


def runProgram(button):
    app.showSubWindow('Обработка')
    addr_file_full_path = app.getEntry('addr_path')
    output_file_full_path = app.getEntry('save_as')
    locality_result = create_locality_list(addr_file_full_path, app)
    create_spreadsheet(locality_result, output_file_full_path)
    

app.setSticky('w')
app.addLabel('addr_path_descr', 'Путь к файлу адресов:', 0, 0, 3)
app.setSticky('ew')
app.addEntry('addr_path', 1, 0, 2)
app.setEntry('addr_path', '{}\\addr.csv'.format(getcwd()))
app.setSticky('w')
app.addIconButton('bt1', open_file, 'open', 1, 2)
app.addEmptyLabel('l1', 2, 0, 3)
app.addLabel('save_as_descr', 'Сохранить в:', 3, 0, 3)
app.setSticky('ew')
app.addEntry('save_as', 4, 0, 2)
app.setEntry('save_as', '{}\\output.xlsx'.format(getcwd()))
app.setSticky('w')
app.addIconButton('bt2', save_file, 'open', 4, 2)
app.setSticky('ew')
app.addEmptyLabel('l2', 5, 0, 3)
app.addButton('Старт', runProgram, 6, 1)
app.addEmptyLabel('l3', 7, 0, 3)
app.addMeter('progress', 8, 0, 3)
app.setMeterFill('progress', 'blue')

app.startSubWindow("Обработка", modal=False)
app.setGeometry("320x220")
app.setPadding([2,0])
app.setInPadding([0,0])
app.setGuiPadding(4,14)
app.setStretch('column')
app.setFont(12)
app.setBg('lightBlue')
app.addLabel('console_window', '')
app.stopSubWindow()

app.go()
