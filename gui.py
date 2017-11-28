from os import getcwd
from appJar import gui
import time
import queue
from coords import *

# class custom_gui(gui):
#     # def __init__(self, title=None, geom=None, warn=None, debug=None, handleArgs=True, language=None, startWindow=None, useTtk=False):
#     #     super().__init__(self, title, geom, warn, debug, handleArgs, language, startWindow, useTtk)
#
#     def

app = gui('Гео', '400x238')
app.setPadding([2,0])
app.setInPadding([0,0])
app.setGuiPadding(4,14)
app.setStretch('column')
font = 'Verdana 12'
app.setBg('lightBlue')
app.setLocation("center")
app.setResizable(False)

result_queue = queue.Queue()
interrupt_queue = queue.Queue()


# class InterruptableThread(threading.Thread):
#     def __init__(self, group=None, target=None, name=None,
#                  args=(), kwargs=None, *, daemon=None):
#         super().__init__(group=group, target=target, name=name,
#                  args=args, kwargs=kwargs, daemon=daemon)
#         self.interrupt_event = threading.Event()
#
#     def run(self):
#         try:
#             if self._target:
#                 self._target(*self._args, **self._kwargs)
#         finally:
#             # Avoid a refcycle if the thread is running a function with
#             # an argument that has a member that points to the thread.
#             del self._target, self._args, self._kwargs


def open_file(button):
    a = app.openBox()
    if a != '':
        app.setEntry('addr_path', a)


def save_file(button):
    a = app.saveBox()
    if a != '':
        app.setEntry('save_as', a)


def press(button):
    p1 = app.getEntry('addr_path')
    p2 = app.getEntry('save_as')
    with open(p1) as f:
        print(f.readline())
    with open(p2) as f:
        print(f.readline())


def stop(button):
    interrupt_queue.put(1)
    app.setGeometry("400x238")
    app.hideButton(button)
    app.showButton('Старт')


def runProgram(button):
    # app.showSubWindow('Обработка')
    app.hideButton('Старт')
    app.showButton('Отмена')
    app.setGeometry("400x280")
    addr_file_full_path = app.getEntry('addr_path')
    output_file_full_path = app.getEntry('save_as')
    # locality_list_thread = InterruptableThread(create_locality_list, addr_file_full_path=addr_file_full_path, app=app, result_queue=result_queue, update_queue=update_queue)
    # locality_list_thread.start()
    app.thread(create_locality_list, addr_file_full_path, app, None, result_queue, interrupt_queue)
    # def update_status_bar():
    #     try:
    #         a = update_queue.get_nowait()
    #     except queue.Empty:
    #         pass
    #     if 'a' in locals():
    #         app.queueFunction(app.setMeter, "progress", a)
    # app.registerEvent(update_status_bar)
    #--
    # while locality_list_thread.is_alive() or not update_queue.empty():
    #     print('alive check')
    #     try:
    #         a = update_queue.get_nowait()
    #     except queue.Empty:
    #         pass
    #     if 'a' in locals():
    #         app.setMeter("progress", a)
    #     time.sleep(.1)
    # locality_list_thread.join()
    # print('after join')

    def save_result():
        try:
            locality_result = result_queue.get_nowait()
        except queue.Empty:
            pass
        if 'locality_result' in locals():
            create_spreadsheet(locality_result, output_file_full_path)
            app.hideButton('Отмена')
            app.showButton('Готово')
    app.registerEvent(save_result)
    # try:
    #     locality_result = result_queue.get_nowait()
    # except queue.Empty:
    #     pass
    # if 'locality_result' in locals():
    #     create_spreadsheet(locality_result, output_file_full_path)





app.setSticky('w')
app.addLabel('addr_path_descr', 'Путь к файлу адресов:', 0, 0, 3)
app.setSticky('ew')
app.addEntry('addr_path', 1, 0, 2)
app.setEntry('addr_path', 'G:\\Py\geo\\addr_short.csv')
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
app.addMeter('progress', 8, 0, 3)
app.setMeterRelief('progress', 'flat')
app.setMeterFill('progress', 'blue')
app.addButton('Отмена', stop, 6, 1)
app.addButton('Готово', stop, 6, 1)
app.setButtonRelief('Отмена', 'groove')


# app.startSubWindow("Обработка", modal=False)
# app.setGeometry("400x100")
# app.setPadding([4,8])
# app.setInPadding([0,0])
# app.setStretch('column')
# app.setFont(12)
# app.setBg('lightBlue')
# app.setLocation("CENTER")
# app.setResizable(False)
# app.hideTitleBar()
#
# app.setSticky('ew')
# app.addMeter('progress', 0, 0, 3)
# app.setMeterRelief('progress', 'flat')
# app.setMeterFill('progress', 'blue')
# app.addButton('Отмена', stop, 1, 1)
# app.setButtonRelief('Отмена', 'groove')
# app.addEmptyLabel('lw1', 2, 2, 3)
#
#
#
# app.stopSubWindow()
app.hideButton('Отмена')
app.hideButton('Готово')
app.go()
