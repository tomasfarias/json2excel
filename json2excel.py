import argparse

from converter import Converter

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Transform a JSON file into Excel or Excel-formatted CSV'
    )
    parser.add_argument('json_file', type=str, help='a json file name or path to json file')
    parser.add_argument(
        '--csv', help='export to CSV instead of Excel (default: False)', action='store_true', default=False
    )
    parser.add_argument(
        '-n', '--name', help='file name for the exported Excel/CSV, without extension (default: result)', type=str, default='result'
    )

    args = parser.parse_args()

    convert = Converter(args.json_file, args.csv, args.name)
    convert.convert()