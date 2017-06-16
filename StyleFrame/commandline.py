import argparse
import json
from collections import defaultdict

import pandas as pd

from StyleFrame import StyleFrame, Container, Styler, version


class CommandLineInterface(object):
    def __init__(self, input_path=None, output_path=None, input_json=None):
        self.input_path = input_path
        self.input_json = input_json
        self.excel_writer = StyleFrame.ExcelWriter(output_path)
        self.col_names_to_width = defaultdict(dict)

    def parse_as_json(self):
        try:
            self._load_from_json()
        except TypeError as ex:
            print('Got the following error:\n{}\nExiting.'.format(ex))
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
        if not isinstance(sheets, list):
            raise TypeError('JSON must contain a list of sheets.')
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
                data[col_name].append(Container(cell['value'], Styler(**(cell.get('style')
                                                                         or col.get('style')
                                                                         or default_cell_style
                                                                         or {})).create_style()))
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
    group.add_argument('--json_path', help='path to json file which defines the Excel file')
    group.add_argument('--json', help='json string which defines the Excel file')
    parser.add_argument('--output_path', help='path of output Excel file, defaults to output.xlsx',
                        default='output.xlsx')

    cli_args = parser.parse_args()

    if not cli_args.version and not any((cli_args.json_path, cli_args.json)):
        parser.error('Either --json_path or --json are required when not using -v.')

    return cli_args


def execute_from_command_line():
    cli_args = get_cli_args()
    if cli_args.version:
        print(version.get_all_versions())
        return
    CommandLineInterface(input_path=cli_args.json_path, input_json=cli_args.json,
                         output_path=cli_args.output_path).parse_as_json()


if __name__ == '__main__':
    execute_from_command_line()
