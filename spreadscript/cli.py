import argparse
import json
import sys

from . import doc_split, version, usage
from .spreadscript import SpreadScript


def read_input(file_name, output_handle):
    """Read the input table."""
    output_handle.write('{}\n'.format(
        json.dumps(SpreadScript(file_name).read_input())))


def read_output(file_name, output_handle):
    """Read the output table."""
    output_handle.write('{}\n'.format(
        json.dumps(SpreadScript(file_name).read_output())))


def process(file_name, data, output_handle):
    """Process the data and read the output table."""
    spreadsheet = SpreadScript(file_name)
    spreadsheet.write_input(json.loads(data))
    output_handle.write('{}\n'.format(json.dumps(spreadsheet.read_output())))


def main():
    """Main entry point."""
    base_parser = argparse.ArgumentParser(add_help=False)
    base_parser.add_argument(
        'file_name', metavar='FILENAME', type=str, help='spreadsheet file')
    base_parser.add_argument(
        '-o', dest='output_handle', metavar='OUTPUT',
        type=argparse.FileType('w'), default=sys.stdout, help='output file')

    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=usage[0], epilog=usage[1])
    parser.add_argument('-v', action='version', version=version(parser.prog))
    subparsers = parser.add_subparsers(dest='subcommand')

    subparser = subparsers.add_parser(
        'read_input', parents=[base_parser],
        description=doc_split(read_input))
    subparser.set_defaults(func=read_input)

    subparser = subparsers.add_parser(
        'read_output', parents=[base_parser],
        description=doc_split(read_output))
    subparser.set_defaults(func=read_output)

    subparser = subparsers.add_parser(
        'process', parents=[base_parser],
        description=doc_split(process))
    subparser.add_argument(
        'data', metavar='DATA', type=str, help='data in JSON format')
    subparser.set_defaults(func=process)

    try:
        args = parser.parse_args()
    except IOError as error:
        parser.error(error)

    if not args.subcommand:
        parser.error('too few arguments')

    try:
        args.func(**{k: v for k, v in vars(args).items()
            if k not in ('func', 'subcommand')})
    except ValueError as error:
        parser.error(error)


if __name__ == '__main__':
    main()
