from os import getcwd
from appJar import gui
import queue
import math
import geocoder
import pickle
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

#####################################
# CORE FUNCTIONS
#####################################

class Location(object):
    def __init__(self,
                 name,
                 latitude,
                 longitude,
                 ):
        self.id = name
        self.lat = latitude
        self.lng = longitude

    def get_name(self):
        return self.id

    def get_lat(self):
        return self.lat

    def get_lng(self):
        return self.lng

    def get_distance(self, other):
        start_lat = math.radians(self.get_lat())
        start_lng = math.radians(self.get_lng())
        end_lat = math.radians(other.get_lat())
        end_lng = math.radians(other.get_lng())
        diff_lat = end_lat - start_lat
        diff_lng = end_lng - start_lng
        a = math.sin(diff_lat / 2) ** 2 + math.cos(start_lat) * math.cos(end_lat) * math.sin(diff_lng / 2) ** 2
        c = 2 * math.asin(math.sqrt(a))
        return round(6371 * c)


class Locality(Location):
    def __init__(self,
                 name,
                 latitude,
                 longitude,
                 county,
                 state,
                 description,
                 ):
            super().__init__(name, latitude, longitude)
            self.county = county
            self.state = state
            self.description = description
            self.adm_center = None
            self.find_adm_center()

    def find_adm_center(self):
        if self.county and 'автономный округ' in self.county.lower():  # если объект в автономном округе, берём его центр вместо областного
            self.adm_center = adm_centers[self.county]
        else:
            self.adm_center = adm_centers[self.state]

    def get_adm_center(self):
        return self.adm_center.get_name()

    def get_adm_center_distance(self):
        return self.get_distance(self.adm_center)

    def find_closest_adm_center(self):
        """
        Iterates over the dict of administrative centers to find the closest
        :return: a tuple, (name, distance)
        """
        distances = {}
        for key in adm_centers:
            distances[self.get_distance(adm_centers[key])] = adm_centers[key].get_name()
        closest = min(distances.keys())
        return (distances[closest], closest)


def create_locality_list(addr_file_full_path, app=None, result_queue=None, interrupt_queue=None):
    """
    Creates a list of Locality objects from the addresses file, stores problematic addresses in a separate list
    :param addr_file_full_path: the full path (including the file name) to the text file with one address on each line
    :return: a tuple, (obj list: Localities, str list: problematic addresses, bool: errors flag)
    """
    errors_flag = False
    locality_list = []
    failed_requests = []
    with open(addr_file_full_path) as f:
        total_addresses = len(f.readlines())
        one_percent = total_addresses/100
        f.seek(0)
        for i,line in enumerate(f):
            if interrupt_queue and not interrupt_queue.empty():
                interrupt_queue.get()
                if app:
                    app.queueFunction(app.setMeter, "progress", 0)
                return
            address = line.strip()
            loc_json = geocoder.yandex(address, lang='ru-RU').json
            try:
                locality = Locality(
                    address,
                    float(loc_json['lat']),
                    float(loc_json['lng']),
                    loc_json.get('county'),  # there may be no county, we use get() to avoid an exception
                    loc_json['state'],
                    loc_json['description']
                )
                locality_list.append(locality)
                if app:
                    app.queueFunction(app.setMeter, "progress", i // one_percent)
                else:
                    print('.', end='', flush=True)
            except (KeyError, TypeError):
                failed_requests.append(address)
                if app:
                    app.queueFunction(app.setMeter, "progress", i // one_percent)
                else:
                    print('X', end='', flush=True)
    print('')
    if app:
        app.queueFunction(app.setMeter, "progress", 100)
    if failed_requests:
        errors_flag = True
        print('Возникли проблемы с частью запросов, проверьте список ошибок')
    if result_queue:
        result_queue.put_nowait((locality_list, failed_requests, errors_flag))
    return (locality_list, failed_requests, errors_flag)


def create_spreadsheet(locality_list_tuple: tuple, output_file_full_path: str):
    """
    :param locality_list_tuple: resulting tuple of the create_locality_list() function
    :param output_file_full_path: full path (including the file name) of the excel spreadsheet in which to save the data
    :return: None, saves an excel table in the current dir
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Адреса'
    ws.append(['Адрес', 'Адм. центр осн.', 'Расстояние осн., км', 'Адм. центр ближ.', 'Расстояние ближ., км'])
    for item in locality_list_tuple[0]:
        if item.get_adm_center() != item.find_closest_adm_center()[0]:
            ws.append([
                item.get_name(),
                item.get_adm_center(),
                item.get_adm_center_distance(),
                item.find_closest_adm_center()[0],
                item.find_closest_adm_center()[1],
            ])
        else:
            ws.append([
                item.get_name(),
                item.get_adm_center(),
                item.get_adm_center_distance(),
                '--',
                '--',
            ])
    tab = Table(displayName='AddrTable', ref='A1:E{}'.format(len(locality_list_tuple[0]) + 1))
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)
    if locality_list_tuple[-1]:
        wb.create_sheet('Ошибки')
        ws = wb['Ошибки']
        ws.append(['Проблемные адреса:'])
        for item in locality_list_tuple[1]:
            ws.append([item])
    wb.save(output_file_full_path)


#####################################
# GUI FUNCTIONS
#####################################

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


def info(button):
    message = '''
Как пользоваться:\n
1. В файле excel удалить все столбцы, кроме столбцов адреса. Адрес может быть разбит на несколько столбцов: например в первом столбце город, а во втором - остальное. Главное, чтобы в каждой строке таблицы не было ничего, кроме адреса.\n
2. Экспортировать таблицу в формат CSV:
Файл -> Экспорт -> Изменить тип файла -> CSV (Разделители - запятые)\n
3. Запустить Geo и нажать на кнопку рядом со строкой "Путь к файлу адресов". В открывшемся окне найти и выбрать созданный файл CSV.\n
4. В строке "Сохранить результат в" указать путь и название файла, в котором будет сохранён результат. Для этого можно также воспользоваться кнопкой обзора.\n
5. Нажать кнопку "Старт" и дождаться окончания обработки. Файл появится в папке, указанной в п. 4\n
'''
    app.infoBox('Инструкция', message, parent=None)


def save_result():
    try:
        locality_result = result_queue.get_nowait()
        create_spreadsheet(locality_result, output_file_full_path)
        app.hideButton('Отмена')
        app.showButton('Готово')
    except queue.Empty:
        pass


def runProgram(button):
    global addr_file_full_path, output_file_full_path
    app.hideButton('Старт')
    app.showButton('Отмена')
    app.setGeometry("400x280")
    addr_file_full_path = app.getEntry('addr_path')
    output_file_full_path = app.getEntry('save_as')
    app.thread(create_locality_list, addr_file_full_path, app, result_queue, interrupt_queue)



#####################################
# RUNNING IN A GUI
#####################################

if __name__ == '__main__':

    with open('adm_centers', 'rb') as file:
        adm_centers = pickle.load(file)

    result_queue = queue.Queue()
    interrupt_queue = queue.Queue()

    # GUI initialization
    app = gui('Geo', '400x238')
    app.setPadding([2, 0])
    app.setInPadding([0, 0])
    app.setGuiPadding(4, 4)
    app.setStretch('column')
    font = 'Verdana 12'
    app.setBg('lightBlue')
    app.setLocation("center")
    app.setResizable(False)
    app.setSticky('w')
    app.addLabel('addr_path_descr', 'Путь к файлу адресов:', 0, 0, 2)
    app.setSticky('ne')
    app.addIconButton('info', info, 'info', 0, 2)
    app.setButtonRelief('info', 'groove')
    app.setSticky('ew')
    app.addEntry('addr_path', 1, 0, 2)
    app.setEntry('addr_path', 'G:\\Py\geo\\addr_short.csv')
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
    app.setButtonRelief('Готово', 'groove')
    app.hideButton('Отмена')
    app.hideButton('Готово')

    app.registerEvent(save_result)
    app.go()





