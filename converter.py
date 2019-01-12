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

    def convert(self, csv_sep=',', helper=None):
        """
        Converts JSON file to either Excel or CSV file, depending on
        self.csv state. File is saved to self.name

        Arguments:
            csv_sep (str): CSV delimiter, for example ',' or ';'.
            helper (dict): See recursive_dict_of_lists.
        """
        flattened = recursive_dict_of_lists(self.file, helper)
        filename = f'{self.name}.{self.export_format}'

        if self.export_format == 'csv':
            with open(filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter=csv_sep)
                writer.writerow(flattened.keys())
                writer.writerows(zip_longest(*flattened.values()))

        elif self.export_format == 'xlsx':
            wb = Workbook()
            ws = wb.active
            ws.append((_ for _ in flattened.keys()))  # doesn't take dict_keys ...
            for row in zip_longest(*flattened.values()):
                    ws.append(row)
            wb.save(filename)

        else:
            raise ValueError(f"Unsupported export_format {self.export_format}. Use 'csv' or 'xlsx'")


def recursive_dict_of_lists(d, helper=None, prev_key=None):
    """
    Builds dictionary of lists by recursively traversing a JSON-like
    structure.

    Arguments:
        d (dict): JSON-like dictionary.
        prev_key (str): Prefix used to create dictionary keys like: prefix_key.
            Passed by recursive step, not intended to be used.
        helper (dict): In case d contains nested dictionaries, you can specify
            a helper dictionary with 'key' and 'value' keys to specify where to
            look for keys and values instead of recursive step. It helps with
            cases like: {'action': {'type': 'step', 'amount': 1}}, by passing
            {'key': 'type', 'value': 'amount'} as a helper you'd get
            {'action_step': [1]} as a result.
    """
    d_o_l = {}

    if helper is not None and helper['key'] in d.keys() and helper['value'] in d.keys():
        if prev_key is not None:
            key = f"{prev_key}_{helper['key']}"
        else:
            key = helper['key']

        if key not in d_o_l.keys():
            d_o_l[key] = []
        d_o_l[key].append(d[helper['value']])

        return d_o_l

    for k, v in d.items():
        if isinstance(v, dict):
            d_o_l.update(recursive_dict_of_lists(v, helper=helper, prev_key=k))
        else:
            if prev_key is not None:
                key = f'{prev_key}_{k}'
            else:
                key = k

            if key not in d_o_l.keys():
                d_o_l[key] = []

            if isinstance(v, list):
                d_o_l[key].extend(v)
            else:
                d_o_l[key].append(v)

    return d_o_l
