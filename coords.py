import os
import math
import geocoder
import pickle
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
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


def create_locality_list(addr_file):
    """
    Creates a list of Locality objects from the addresses file, stores problematic addresses in a separate list
    :param addr_file: a text file with one address on each line
    :return: a tuple, (obj list: Localities, str list: problematic addresses, bool: errors flag)
    """
    errors_flag = False
    locality_list = []
    failed_requests = []
    with open(addr_file) as f:
        for line in f:
            address = line.strip()
            loc_json = geocoder.yandex(address, lang='ru-RU').json
            # if loc_json is not None and loc_json['status'] == 'OK':
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
                print('.', end='')
            except (KeyError, TypeError):
                failed_requests.append(address)
                print('X', end='')
            # else:
            #     failed_requests.append(address)
            #     print('X', end='')
    print('')
    if failed_requests:
        errors_flag = True
        print('Возникли проблемы с частью запросов, проверьте список ошибок')
    return (locality_list, failed_requests, errors_flag)


def create_spreadsheet(locality_list_tuple):
    """
    :param locality_list_tuple: resulting tuple of the create_locality_list() function
    :return: saves an excel table in the current dir
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
    wb.save("table_full.xlsx")





if __name__ == '__main__':
    with open('adm_centers', 'rb') as file:
        adm_centers = pickle.load(file)
    locality_result = create_locality_list('addr_full.csv')
    create_spreadsheet(locality_result)
    # for item in locality_result[0]:
    #     print('{}\n{}, {}\nAdministrative center:\n{} - {}km\n{}, {}\nClosest center:\n{} - {}km\n\n'.format(
    #         item.get_name(),
    #         item.get_lat(),
    #         item.get_lng(),
    #         item.get_adm_center(),
    #         item.get_adm_center_distance(),
    #         item.adm_center.get_lat(),  # ! need method or delete
    #         item.adm_center.get_lng(),  # ! need method or delete
    #         item.find_closest_adm_center()[0],
    #         item.find_closest_adm_center()[1],
    #         )
    #     )
    # if locality_result[-1]:
    #     print('Problematic addresses:')
    #     for item in locality_result[1]:
    #         print(item)
