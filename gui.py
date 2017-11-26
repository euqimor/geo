from os import getcwd
from appJar import gui
import time
from queue import Queue
from coords import *

app = gui('Гео', '400x240')
app.setPadding([2,0])
app.setInPadding([0,0])
app.setGuiPadding(4,14)
app.setStretch('column')
font = 'Verdana 12'
app.setBg('lightBlue')
app.setLocation("center")
app.setResizable(False)


class InterruptableThread(threading.Thread):
    def __init__(self, function, **kwargs):
        super().__init__()
        self.interrupt_event = threading.Event()
        self.kwargs = kwargs
        self.function = function

    def run(self):
        print('thread runs')
        self.function(thread_=self, **self.kwargs)
        # function(self, *self.args)


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
    queue_ = Queue()
    locality_list_thread = InterruptableThread(create_locality_list, addr_file_full_path=addr_file_full_path, app=app, queue_=queue_)
    locality_list_thread.start()
    locality_list_thread.join()
    print('after join')
    # while locality_list_thread.is_alive():
    #     print('alive check')
    #     time.sleep(1)
    locality_result = queue_.get()
    if locality_result:
        create_spreadsheet(locality_result, output_file_full_path)


def stop(button):
    # thread_.interrupt_event.set()
    pass


app.setSticky('w')
app.addLabel('addr_path_descr', 'Путь к файлу адресов:', 0, 0, 3)
app.setSticky('ew')
app.addEntry('addr_path', 1, 0, 2)
app.setEntry('addr_path', 'G:\\Py\geo\\trash\\addr_short.csv')
# app.setEntry('addr_path', '{}\\addr.csv'.format(getcwd()))
app.setSticky('w')
app.addIconButton('bt1', open_file, 'open', 1, 2)
app.setButtonRelief('bt1', 'groove')
app.addEmptyLabel('l1', 2, 0, 3)
app.addLabel('save_as_descr', 'Сохранить результат в:', 3, 0, 3)
app.setSticky('ew')
app.addEntry('save_as', 4, 0, 2)
app.setEntry('save_as', '{}\\output.xlsx'.format(getcwd()))
app.setSticky('w')
app.addIconButton('bt2', save_file, 'open', 4, 2)
app.setButtonRelief('bt2', 'groove')
app.setSticky('ew')
app.addEmptyLabel('l2', 5, 0, 3)
app.addButton('Старт', runProgram, 6, 1)
app.setButtonRelief('Старт', 'groove')
app.addEmptyLabel('l3', 7, 0, 3)
# app.addMeter('progress', 8, 0, 3)
# app.setMeterRelief('progress', 'flat')
# app.setMeterFill('progress', 'blue')


app.startSubWindow("Обработка", modal=False)
app.setGeometry("400x100")
app.setPadding([4,8])
app.setInPadding([0,0])
app.setStretch('column')
app.setFont(12)
app.setBg('lightBlue')
app.setLocation("CENTER")
# app.setResizable(False)
app.hideTitleBar()

app.setSticky('ew')
app.addMeter('progress', 0, 0, 3)
app.setMeterRelief('progress', 'flat')
app.setMeterFill('progress', 'blue')
app.addButton('Отмена', stop, 1, 1)
app.setButtonRelief('Отмена', 'groove')
app.addEmptyLabel('lw1', 2, 2, 3)



app.stopSubWindow()

app.go()
