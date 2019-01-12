import json
import csv
from pathlib import Path
from itertools import zip_longest

from openpyxl import Workbook


class Converter:
    """
        Converts JSON file to Excel or CSV file. Nested structures are flattened
        by recursively appending lists of values.
    """

    def __init__(self, file=None, export_format='csv', name='result'):
        if file is not None:
            self.file = file
        self.export_format = export_format
        self.name = name

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        filename = f'{self.name}.{self.export_format}'
        p = Path(filename)

        if exc_type is not None:
            try:
                p.unlink()
            except FileNotFoundError:
                pass
            return False

    @property
    def file(self):
        return self._file

    @file.setter
    def file(self, file):
        """
        Setting property handles reading JSON file into instance attribute.

        Arguments:
            file (str): file name or path to file. Will be used to initialize
                a pathlib.Path object.
        """
        p = Path(file)
        if p.suffix == '':
            p = p.with_suffix('.json')

        with p.open(encoding='utf-8') as f:
            self._file = json.load(f)

    def convert(self, csv_sep=','):
        """
        Converts JSON file to either Excel or CSV file, depending on
        self.csv state. File is saved to self.name

        Arguments:
            csv_sep (str): CSV delimiter, for example ',' or ';'.
        """
        long_list = zip_longest(*recursive_list_of_lists(self.file))
        filename = f'{self.name}.{self.export_format}'

        if self.export_format == 'csv':
            with open(filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter=csv_sep)
                for l in long_list:
                    writer.writerow(l)

        elif self.export_format == 'xlsx':
            wb = Workbook()
            ws = wb.active
            for row in long_list:
                    ws.append(row)
            wb.save(filename)

        else:
            raise ValueError(f"Unsupported export_format {self.export_format}. Use 'csv' or 'xlsx'")


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
