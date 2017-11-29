from os import getcwd
from appJar import gui
import queue
import math
import geocoder
import os
import sys
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
    try:
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
    except FileNotFoundError as e:
        if app:
            app.errorBox('Файл не найден', 'Файл адресов не найден, проверьте правильность указанного пути')
            return
        else:
            raise e
    print('')
    if app:
        app.queueFunction(app.setMeter, "progress", 100)
    if failed_requests:
        errors_flag = True
    if result_queue:
        result_queue.put_nowait((locality_list, failed_requests, errors_flag))
    return (locality_list, failed_requests, errors_flag)


def create_spreadsheet(locality_list_tuple: tuple, output_file_full_path: str, app=None):
    """
    :param locality_list_tuple: resulting tuple of the create_locality_list() function
    :param output_file_full_path: full path (including the file name) of the excel spreadsheet in which to save the data
    :return: None, saves an excel table in the current dir
    """
    try:
        with open(output_file_full_path, 'w') as f:
            f.write('test')
    except FileNotFoundError as e:
        if app:
            app.errorBox('Ошибка', 'Не удаётся создать файл в указанной директории. Убедитесь, что директория существует, и повторите попытку.')
            return
        else:
            raise e
    except PermissionError as e:
        if app:
            app.errorBox('Ошибка', 'Не хватает разрешений для создания файла.')
            return
        else:
            raise e
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
    try:
        wb.save(output_file_full_path)
        if locality_list_tuple[-1] and app:
            app.warningBox('Готово', 'Обработка завершена с ошибками, проверьте лист "Ошибки" в таблице.')
        else:
            if app:
                app.infoBox('Готово', 'Обработка завершена.')
    except Exception as e:
        if app:
            app.errorBox('Ошибка', 'Ошибка при сохранении файла:\n{}'.format(e))
        else:
            raise e


def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)


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
        create_spreadsheet(locality_result, output_file_full_path, app)
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

    adm_centers = {'Республика Адыгея': Location('Майкоп', 44.609764, 40.100516),
                   'Республика Алтай': Location('Горно-Алтайск', 51.958182, 85.960373),
                   'Республика Башкортостан': Location('Уфа', 54.735147, 55.958727),
                   'Республика Бурятия': Location('Улан-Удэ', 51.834464, 107.584574),
                   'Республика Дагестан': Location('Махачкала', 42.98306, 47.504682),
                   'Республика Ингушетия': Location('Магас', 43.166669, 44.80484),
                   'Кабардино-Балкарская Республика': Location('Нальчик', 43.485259, 43.607072),
                   'Республика Калмыкия': Location('Элиста', 46.308309, 44.270181),
                   'Карачаево-Черкесская Республика': Location('Черкесск', 44.226863, 42.04677),
                   'Республика Карелия': Location('Петрозаводск', 61.789036, 34.359688),
                   'Республика Коми': Location('Сыктывкар', 61.668831, 50.836461),
                   'Республика Крым': Location('Симферополь', 44.948314, 34.100192),
                   'Республика Марий Эл': Location('Йошкар-Ола', 56.634407, 47.899878),
                   'Республика Мордовия': Location('Саранск', 54.187211, 45.183642),
                   'Республика Саха (Якутия)': Location('Якутск', 62.028103, 129.732663),
                   'Республика Северная Осетия — Алания': Location('Владикавказ', 43.020603, 44.681888),
                   'Республика Татарстан': Location('Казань', 55.798551, 49.106324),
                   'Республика Тыва': Location('Кызыл', 51.719086, 94.437757),
                   'Удмуртская Республика': Location('Ижевск', 56.852593, 53.204843),
                   'Республика Хакасия': Location('Абакан', 53.721152, 91.442387),
                   'Чеченская Республика': Location('Грозный', 43.317776, 45.694909),
                   'Чувашская Республика': Location('Чебоксары', 56.146247, 47.250153),
                   'Алтайский край': Location('Барнаул', 53.355084, 83.769948),
                   'Забайкальский край': Location('Чита', 52.033973, 113.499432),
                   'Камчатский край': Location('Петропавловск-Камчатский', 53.03704, 158.655918),
                   'Краснодарский край': Location('Краснодар', 45.035566, 38.974711),
                   'Красноярский край': Location('Красноярск', 56.010563, 92.852572),
                   'Пермский край': Location('Пермь', 58.010374, 56.229398),
                   'Приморский край': Location('Владивосток', 43.115141, 131.885341),
                   'Ставропольский край': Location('Ставрополь', 45.044521, 41.969083),
                   'Хабаровский край': Location('Хабаровск', 48.480223, 135.071917),
                   'Амурская область': Location('Благовещенск', 50.290658, 127.527173),
                   'Архангельская область': Location('Архангельск', 64.539393, 40.516939),
                   'Астраханская область': Location('Астрахань', 46.347869, 48.033574),
                   'Белгородская область': Location('Белгород', 50.59566, 36.587223),
                   'Владимирская область': Location('Владимир', 56.129042, 40.40703),
                   'Волгоградская область': Location('Волгоград', 48.707103, 44.516939),
                   'Вологодская область': Location('Вологда', 59.220473, 39.891559),
                   'Воронежская область': Location('Воронеж', 51.661535, 39.200287),
                   'Ивановская область': Location('Иваново', 57.000348, 40.973921),
                   'Иркутская область': Location('Иркутск', 52.286387, 104.28066),
                   'Калининградская область': Location('Калининград', 54.70739, 20.507307),
                   'Калужская область': Location('Калуга', 54.513845, 36.261215),
                   'Кемеровская область': Location('Кемерово', 55.354968, 86.087314),
                   'Кировская область': Location('Киров', 58.603581, 49.667978),
                   'Костромская область': Location('Кострома', 57.767961, 40.926858),
                   'Курганская область': Location('Курган', 55.441606, 65.344316),
                   'Курская область': Location('Курск', 51.730361, 36.192647),
                   'Ленинградская область': Location('Санкт-Петербург', 59.939095, 30.315868),
                   'Липецкая область': Location('Липецк', 52.608782, 39.599346),
                   'Магаданская область': Location('Магадан', 59.568164, 150.808541),
                   'Московская область': Location('Москва', 55.753215, 37.622504),
                   'Мурманская область': Location('Мурманск', 68.969582, 33.074558),
                   'Нижегородская область': Location('Нижний Новгород', 56.326887, 44.005986),
                   'Новгородская область': Location('Великий Новгород', 58.52281, 31.269915),
                   'Новосибирская область': Location('Новосибирск', 55.030199, 82.92043),
                   'Омская область': Location('Омск', 54.989342, 73.368212),
                   'Оренбургская область': Location('Оренбург', 51.768199, 55.096955),
                   'Орловская область': Location('Орёл', 52.970143, 36.063397),
                   'Пензенская область': Location('Пенза', 53.195063, 45.018316),
                   'Псковская область': Location('Псков', 57.819365, 28.331786),
                   'Ростовская область': Location('Ростов-на-Дону', 47.222555, 39.718678),
                   'Рязанская область': Location('Рязань', 54.629148, 39.734928),
                   'Самарская область': Location('Самара', 53.195538, 50.101783),
                   'Саратовская область': Location('Саратов', 51.533103, 46.034158),
                   'Сахалинская область': Location('Южно-Сахалинск', 46.959179, 142.738041),
                   'Свердловская область': Location('Екатеринбург', 56.838011, 60.597465),
                   'Смоленская область': Location('Смоленск', 54.78264, 32.045134),
                   'Тамбовская область': Location('Тамбов', 52.721219, 41.452274),
                   'Тверская область': Location('Тверь', 56.859611, 35.911896),
                   'Томская область': Location('Томск', 56.48466, 84.948179),
                   'Тульская область': Location('Тула', 54.193033, 37.617752),
                   'Тюменская область': Location('Тюмень', 57.153033, 65.534328),
                   'Ульяновская область': Location('Ульяновск', 54.316855, 48.402557),
                   'Челябинская область': Location('Челябинск', 55.160026, 61.40259),
                   'Ярославская область': Location('Ярославль', 57.626569, 39.893787),
                   'Еврейская автономная область': Location('Биробиджан', 48.794662, 132.921736),
                   'Ненецкий автономный округ': Location('Нарьян-Мар', 67.63805, 53.006926),
                   'Ханты-Мансийский автономный округ — Югра': Location('Ханты-Мансийск', 61.00318, 69.018902),
                   'Чукотский автономный округ': Location('Анадырь', 64.734816, 177.514745),
                   'Ямало-Ненецкий автономный округ': Location('Салехард', 66.530715, 66.613851),
                   'Брянская область': Location('Брянск', 53.243325, 34.363731),
                   'Ханты-Мансийский автономный округ': Location('Ханты-Мансийск', 61.00318, 69.018902),
                   'Москва': Location('Москва', 55.753215, 37.622504),
                   'Севастополь': Location('Севастополь', 44.616687, 33.525432),
                   'Санкт-Петербург': Location('Санкт-Петербург', 59.939095, 30.315868)}

    result_queue = queue.Queue()
    interrupt_queue = queue.Queue()

    img_open = resource_path("open.png")
    img_info = resource_path("info.png")
    ico = resource_path("globe2.ico")

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
    # app.addIconButton('info', info, 'info', 0, 2)
    app.addImageButton('info', info, img_info, 0, 2)
    app.setButtonRelief('info', 'groove')
    app.setSticky('ew')
    app.addEntry('addr_path', 1, 0, 2)
    app.setEntry('addr_path', 'G:\\Py\geo\\addr_short.csv')
    app.setSticky('w')
    # app.addIconButton('bt1', open_file, 'open', 1, 2)
    app.addImageButton('bt1', open_file, img_open, 1, 2)
    app.setButtonRelief('bt1', 'groove')
    app.addEmptyLabel('l1', 2, 0, 3)
    app.addLabel('save_as_descr', 'Сохранить результат в:', 3, 0, 3)
    app.setSticky('ew')
    app.addEntry('save_as', 4, 0, 2)
    app.setEntry('save_as', '{}\\output.xlsx'.format(getcwd()))
    app.setSticky('w')
    # app.addIconButton('bt2', save_file, 'open', 4, 2)
    app.addImageButton('bt2', save_file, img_open, 4, 2)
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
    app.setIcon(ico)

    app.registerEvent(save_result)
    app.go()





