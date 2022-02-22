import argparse
import json
import jsonschema
import inspect
import pandas as pd

from collections import defaultdict
from pprint import pprint

from .. import StyleFrame, Container, Styler, version, tests
from .tests.json_schema import commandline_json_schema

styler_kwargs = set(inspect.signature(Styler).parameters.keys())


class CommandLineInterface:
    def __init__(self, input_path=None, output_path=None, input_json=None):
        self.input_path = input_path
        self.input_json = input_json
        self.excel_writer = StyleFrame.ExcelWriter(output_path)
        self.col_names_to_width = defaultdict(dict)

    def parse_as_json(self):
        try:
            self._load_from_json()
        except (TypeError, ValueError) as ex:
            print('Got the following error:\n{}.'.format(ex))
            return
        self._save()

    def _load_from_json(self):
        if self.input_json:
            sheets = json.loads(self.input_json)
        elif self.input_path:
            with open(self.input_path) as f:
                sheets = json.load(f)
        else:
            raise TypeError('Neither --json nor --json_path were provided.')

        try:
            jsonschema.validate(sheets, commandline_json_schema)
        except jsonschema.ValidationError as validation_error:
            raise ValueError(validation_error)

        for sheet in sheets:
            self._load_sheet(sheet)

    def _load_sheet(self, sheet):
        sheet_name = sheet['sheet_name']
        default_cell_style = sheet.get('default_styles', {}).get('cells')
        data = defaultdict(list)
        for col in sheet['columns']:
            col_name = col['col_name']
            col_width = col.get('width')
            if col_width:
                self.col_names_to_width[sheet_name][col_name] = col_width
            for cell in col['cells']:
                provided_style = cell.get('style') or col.get('style') or default_cell_style or {}
                unrecognized_styler_kwargs = set(provided_style.keys()) - styler_kwargs
                if unrecognized_styler_kwargs:
                    raise TypeError('Styler dict {} contains unexpected argument: {}.\n'
                                    'Expected arguments: {}'.format(provided_style, unrecognized_styler_kwargs,
                                                                    styler_kwargs))
                else:
                    data[col_name].append(Container(cell['value'], Styler(**provided_style)))
        sf = StyleFrame(pd.DataFrame(data=data))

        self._apply_headers_style(sf, sheet)
        self._apply_cols_and_rows_dimensions(sf, sheet)
        sf.to_excel(excel_writer=self.excel_writer, sheet_name=sheet_name, **sheet.get('extra_features', {}))
        setattr(self, '{}_sf'.format(sheet_name), sf)

    def _apply_headers_style(self, sf, sheet):
        default_headers_style = sheet.get('default_styles', {}).get('headers')
        if default_headers_style:
            sf.apply_headers_style(styler_obj=Styler(**default_headers_style))

    def _apply_cols_and_rows_dimensions(self, sf, sheet):
        sf.set_column_width_dict(self.col_names_to_width[sheet['sheet_name']])
        row_heights = sheet.get('row_heights')
        if row_heights:
            sf.set_row_height_dict(row_heights)

    def _save(self):
        self.excel_writer.save()


def get_cli_args():
    parser = argparse.ArgumentParser('Command-line interface for StyleFrame library')
    group = parser.add_mutually_exclusive_group()

    group.add_argument('-v', '--version', action='store_true', default=False,
                       help='print versions of the Python interpreter, openpyxl, pandas and StyleFrame then quit')
    group.add_argument('--json_path', '--json-path', help='path to json file which defines the Excel file')
    group.add_argument('--json', help='json string which defines the Excel file')
    group.add_argument('--show-schema', action='store_true', help='Print the JSON schema used for validation and exit',
                       default=False)
    group.add_argument('--test', help='execute tests', action='store_true')
    parser.add_argument('--output_path', '--output-path', help='path of output Excel file, defaults to output.xlsx',
                        default='output.xlsx')

    cli_args = parser.parse_args()

    if not any((cli_args.version, cli_args.show_schema, cli_args.test)) and not any((cli_args.json_path, cli_args.json)):
        parser.error('Either --json_path or --json are required when not using -v or --show-schema')

    return cli_args


def execute_from_command_line():
    cli_args = get_cli_args()
    if cli_args.version:
        print(version.get_all_versions())
        return
    if cli_args.show_schema:
        pprint(commandline_json_schema)
        return
    if cli_args.test:
        tests.tests.run()
        return
    CommandLineInterface(input_path=cli_args.json_path, input_json=cli_args.json,
                         output_path=cli_args.output_path).parse_as_json()


if __name__ == '__main__':
    execute_from_command_line()
