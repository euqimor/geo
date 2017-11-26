import os
import math
import geocoder
import pickle
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import threading
import sys


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


def create_locality_list(addr_file_full_path, app=None, thread_=None, result_queue=None, update_queue=None):
    """
    Creates a list of Locality objects from the addresses file, stores problematic addresses in a separate list
    :param addr_file_full_path: the full path (including the file name) to the text file with one address on each line
    :return: a tuple, (obj list: Localities, str list: problematic addresses, bool: errors flag)
    """
    errors_flag = False
    locality_list = []
    failed_requests = []
    lock = threading.Lock()
    with open(addr_file_full_path) as f:
        total_addresses = len(f.readlines())
        one_percent = total_addresses/100
        f.seek(0)
        for i,line in enumerate(f):
            if thread_ is None or (thread_ and not thread_.interrupt_event.is_set()):
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
                    if update_queue:
                        update_queue.put_nowait(i // one_percent)
                    # if app:
                    #     with lock:
                    #         app.setMeter("progress", i//one_percent)
                            # app.setMessage('console_window', app.getMessage('console_window') + '.')
                    else:
                        print('.', end='', flush=True)
                except (KeyError, TypeError):
                    failed_requests.append(address)
                    if update_queue:
                        update_queue.put_nowait(i // one_percent)
                    # if app:
                    #     with lock:
                    #         app.setMeter("progress", i // one_percent)
                            # app.setMessage('console_window', app.getMessage('console_window') + 'X')
                    else:
                        print('X', end='', flush=True)
            else:
                print('Stopping due to an interrupt')
                exit(0)
    print('')
    if update_queue:
        update_queue.put_nowait(100)
    # if app:
    #     with lock:
    #         app.setMeter("progress", 100)
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


with open('adm_centers', 'rb') as file:
    adm_centers = pickle.load(file)
