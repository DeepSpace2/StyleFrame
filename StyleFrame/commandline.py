"""Command-line interface for StyleFrame library.

Usage:
    styleframe <json_file_path> [--output_path=<path>]
    styleframe (-h | --help | --version)

Options:
    --output_path=<path> The path of the output file [default: output.xlsx]

"""
import docopt
from StyleFrame import version
import json
import pandas as pd
from StyleFrame import StyleFrame, Container, Styler


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

        df = pd.DataFrame(data=data, columns=[column['value'] for column in sheet['sheet_columns']])
        sf = StyleFrame(df)

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


def execute_from_command_line():
    argv = docopt.docopt(doc=__doc__, version=version._version_)
    CommandLineInterface(input_path=argv['<json_file_path>'], output_path=argv['--output_path']).parse_as_json()

if __name__ == '__main__':
    execute_from_command_line()