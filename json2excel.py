import argparse

from converter import Converter

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Transform a JSON file into Excel or Excel-formatted CSV'
    )
    parser.add_argument(
        '--file', type=str, help='a json or json-like file path'
    )
    parser.add_argument(
        '--format', help='export format: csv or xlsx (default: csv)', type=str, default='csv'
    )
    parser.add_argument(
        '-n', '--name', help='file name for the exported Excel/CSV, without extension (default: result)', type=str, default='result'
    )
    parser.add_argument(
        '-s', '--sep', help='separator to use for the CSV file (default: ,)', type=str, default=','
    )

    args = parser.parse_args()

    with Converter(args.file, args.format, args.name) as c:
        c.convert(csv_sep=args.sep)
