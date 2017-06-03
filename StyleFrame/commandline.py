"""Command-line interface for StyleFrame library.

Usage:
    styleframe <json_file_path> [--output_path=<path>]
    styleframe (-h | --help | --version)

Options:
    --output_path=<path> The path of the output file [default: output.xlsx]

"""
import docopt
import version
import json
import pandas as pd
from style_frame import StyleFrame, Container, Styler


class CommandLineInterface(object):
    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.excel_writer = StyleFrame.ExcelWriter(output_path)

    def parse_as_json(self):
        self._load_from_json()
        self._save()

    def _load_from_json(self):
        with open(self.input_path) as f:
            input_json = json.load(f)

            if not isinstance(input_json, list):
                input_json = [input_json]

            for sheet in input_json:
                self._load_single_sheet(sheet)

    def _load_single_sheet(self, sheet):
        default_style = Styler(**sheet.get('default_style', {}))
        data = {col_name: [Container(item['value'], Styler(**item['style']).create_style())
                           if 'style' in item else Container(item['value'], default_style.create_style())
                           for item in items]
                for col_name, items in sheet['sheet_data'].iteritems()}

        df = pd.DataFrame(data=data, columns=[item['value'] for item in sheet['sheet_columns']])
        sf = StyleFrame(df)

        sf.to_excel(excel_writer=self.excel_writer,
                    sheet_name=sheet['sheet_name'],
                    **sheet['extra_features'])

    def _save(self):
        self.excel_writer.save()


if __name__ == "__main__":
    argv = docopt.docopt(doc=__doc__, version=version._version_)
    print argv
    CommandLineInterface(input_path=argv['<json_file_path>'], output_path=argv['--output_path']).parse_as_json()
