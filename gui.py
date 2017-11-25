from os import getcwd
from appJar import gui

app = gui('Гео', '320x160')
app.setPadding([2,0])
app.setInPadding([0,0])
app.setGuiPadding(4,4)
app.setStretch('column')
app.setFont(12)

def press(button):
    if button == "Cancel":
        app.stop()
    else:
        usr = app.getEntry("Username")
        pwd = app.getEntry("Password")
        print("User:", usr, "Pass:", pwd)

app.setSticky("w")
app.addLabel("path_descr", "Путь к файлу адресов:", 0, 0, 3)
app.setSticky("ew")
app.addEntry("path", 1, 0, 2)
app.setEntry('path', '{}\\addr.csv'.format(getcwd()))
app.setSticky("w")
app.addIconButton('Обзор', press, 'open', 1, 2)
app.addLabel("name_descr", "Название итоговой таблицы:", 2, 0, 3)
app.setSticky("ew")
app.addEntry("name", 3, 0, 2)
app.setEntry('name', 'output.xlsx')
app.addEmptyLabel('l1', 4, 0, 3)
app.addButton('Старт', press, 5, 1)
# app.addMeter("progress", 5, 0, 3)
# app.setMeterFill("progress", "blue")

app.go()
