import argparse
import json
import pandas as pd

from StyleFrame import StyleFrame, Container, Styler, version


class CommandLineInterface(object):
    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.excel_writer = StyleFrame.ExcelWriter(output_path)

    def parse_as_json(self):
        self._load_from_json()
        self._save()

    def _load_from_json(self):
        with open(self.input_path) as j:
            sheets = json.load(j)

            if not isinstance(sheets, list):
                sheets = list(sheets)

            for sheet in sheets:
                self._load_single_sheet(sheet)

    def _load_single_sheet(self, sheet):
        default_style = Styler(**sheet.get('default_style', {})).create_style()
        data = {col_name: [Container(cell['value'],
                                     Styler(**cell['style']).create_style() if 'style' in cell else default_style)
                           for cell in cells]
                for col_name, cells in sheet['sheet_data'].items()}

        sf = StyleFrame(pd.DataFrame(data=data, columns=[column['value'] for column in sheet['sheet_columns']]))

        if 'default_columns_style' in sheet:
            default_columns_style = sheet['default_columns_style']
            if 'width' in default_columns_style:
                sf.set_column_width(columns=list(sf.columns), width=default_columns_style.pop('width'))

            if default_columns_style:
                style = Styler(**default_columns_style)
                sf.apply_column_style(cols_to_style=list(sf.columns), styler_obj=style)

        if 'default_header_style' in sheet:
            sf.apply_headers_style(styler_obj=Styler(**sheet['default_header_style']))

        for column in filter(lambda col: 'width' in col, sheet['sheet_columns']):
            sf.set_column_width(columns=column['value'], width=column['width'])

        sf.to_excel(excel_writer=self.excel_writer,
                    sheet_name=sheet['sheet_name'],
                    **sheet['extra_features'])

    def _save(self):
        self.excel_writer.save()


def get_cli_args():
    parser = argparse.ArgumentParser('Command-line interface for StyleFrame library')
    group = parser.add_mutually_exclusive_group()

    group.add_argument('-v', '--version', action='store_true', default=False,
                       help='print versions of the Python interpreter, openpyxl, pandas and StyleFrame then quit')
    group.add_argument('--json_path', help='path to json file which defines the Excel file')

    parser.add_argument('--output_path', help='path of output Excel file, defaults to output.xlsx', default='output.xlsx')

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
