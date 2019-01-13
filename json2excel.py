import argparse

from converter import Converter

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Transform a JSON file into Excel or Excel-formatted CSV'
    )
    parser.add_argument(
        '-i', '--input-file', type=str, help='a json or json-like input file path'
    )
    parser.add_argument(
        '-o', '--ouput-file', type=str, default='result.csv',
        help="output file path, format is inferred by suffix (.csv or .xlsx) (default: 'result.csv')",
    )
    parser.add_argument(
        '-s', '--separator', type=str, default=',',
        help='separator to use for the CSV file (default: ,)'
    )

    args = parser.parse_args()

    with Converter(args['input-file'], args['ouput-file']) as c:
        c.convert(csv_sep=args.separator)
