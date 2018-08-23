import json
import csv
from pathlib import Path
from itertools import zip_longest

from openpyxl import Workbook

class Converter:

    def __init__(self, json_file=None, csv=False, name='result'):
        if json_file is not None:
            self.json_file = json_file
        self.csv = csv
        self.name = name
    
    @property
    def json_file(self):
        return self._json_file

    @json_file.setter
    def json_file(self, file):
        p = Path(file)
        if p.suffix != '.json':
            p = p.with_suffix('.json')
        with p.open(encoding='utf-8') as f:
            self._json_file = json.load(f)

    def convert(self):
        long_list = zip_longest(*recursive_list_of_lists(self.json_file))
        if self.csv == True:
            filename = f'{self.name}.csv'
            with open(filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                for l in long_list:
                    writer.writerow(l)
        else:
            filename = f'{self.name}.xlsx'
            wb = Workbook()
            ws = wb.active
            for row in long_list:
                    ws.append(row)
            wb.save(filename)

def recursive_keys(d):
    keys = []
    for k in d.keys():
        if isinstance(d[k], dict):
            keys += recursive_keys(d[k])
        else:
            keys.append(k)
    return keys

def recursive_values(d):
    values = []
    for k in d.keys():
        if isinstance(d[k], dict):
            values += recursive_values(d[k])
        else:
            values.append(d[k])
    return values

def recursive_dict(d, prev_key=None):
    dic = {}
    for k in d.keys():
        if isinstance(d[k], dict):
            dic.update(recursive_dict(d[k], prev_key=k))
        else:
            if prev_key is not None:
                new_key = f'{prev_key}_{k}'
            else:
                new_key = k
            dic[new_key] = d[k]
    return dic

def recursive_list_of_lists(d, prev_key=None):
    list_ = []
    for k in d.keys():
        if isinstance(d[k], dict):
            list_.extend(recursive_list_of_lists(d[k], prev_key=k))
        else:
            if prev_key is not None:
                new_list = [f'{prev_key}_{k}']
            else:
                new_list = [k]
            if isinstance(d[k], list):
                for v in d[k]:
                    new_list.append(v)
            else:
                new_list.append(d[k])
            list_.append(new_list)
    return list_
