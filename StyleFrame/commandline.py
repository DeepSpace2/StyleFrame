import argparse
import json
from collections import defaultdict

import pandas as pd

from StyleFrame import StyleFrame, Container, Styler, version


class CommandLineInterface(object):
    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.excel_writer = StyleFrame.ExcelWriter(output_path)
        self.col_names_to_width = defaultdict(dict)

    def parse_as_json(self):
        self._load_from_json()
        self._save()

    def _load_from_json(self):
        with open(self.input_path) as j:
            sheets = json.load(j)
            if not isinstance(sheets, list):
                sheets = list(sheets)
            for sheet in sheets:
                self._load_sheet(sheet)

    def _load_sheet(self, sheet):
        sheet_name = sheet['sheet_name']
        data = defaultdict(list)
        for col in sheet['columns']:
            col_name = col['col_name']
            col_width = col.get('width')
            if col_width:
                self.col_names_to_width[sheet_name][col_name] = col_width
            for cell in col['cells']:
                data[col_name].append(Container(cell['value'], Styler(**(cell.get('style')
                                                                         or col.get('style')
                                                                         or {})).create_style()))
        sf = StyleFrame(pd.DataFrame(data=data))

        self._apply_headers_style(sf, sheet)
        self._apply_cols_and_rows_dimensions(sf, sheet)
        sf.to_excel(excel_writer=self.excel_writer, sheet_name=sheet_name, **sheet['extra_features'])
        setattr(self, '{}_sf'.format(sheet_name), sf)

    def _apply_headers_style(self, sf, sheet):
        sf.apply_headers_style(styler_obj=Styler(**(sheet.get('default_styles', {}).get('headers') or {})))

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

    parser.add_argument('--output_path', help='path of output Excel file, defaults to output.xlsx',
                        default='output.xlsx')

    cli_args = parser.parse_args()

    if not cli_args.version and not cli_args.json_path:
        parser.error('--json_path is required when not using -v.')

    return cli_args


def execute_from_command_line():
    cli_args = get_cli_args()
    if cli_args.version:
        print(version.get_all_versions())
        return
    CommandLineInterface(input_path=cli_args.json_path, output_path=cli_args.output_path).parse_as_json()


if __name__ == '__main__':
    execute_from_command_line()
