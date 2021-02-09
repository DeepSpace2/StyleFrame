import argparse
import unittest
import os

from contextlib import suppress
from unittest.mock import patch

from styleframe import CommandLineInterface, Styler, utils
from styleframe.command_line.commandline import get_cli_args
from styleframe.command_line.tests import TEST_JSON_FILE, TEST_JSON_STRING_FILE
from styleframe.tests import TEST_FILENAME


class CommandlineInterfaceTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sheet_1_col_a_style = Styler(bg_color=utils.colors.blue, font_color=utils.colors.yellow).to_openpyxl_style()
        cls.sheet_1_col_a_cell_2_style = Styler(bold=True, font=utils.fonts.arial, font_size=30,
                                                font_color=utils.colors.green,
                                                border_type=utils.borders.double).to_openpyxl_style()
        cls.sheet_1_col_b_cell_4_style = Styler(bold=True, font=utils.fonts.arial, font_size=16).to_openpyxl_style()

    def tearDown(self):
        with suppress(OSError):
            os.remove(TEST_FILENAME)

    # noinspection PyUnresolvedReferences
    def test_parse_as_json(self):
        cli = CommandLineInterface(TEST_JSON_FILE, TEST_FILENAME)
        cli.parse_as_json()
        loc_col_a = cli.Sheet1_sf.columns.get_loc('col_a')
        loc_col_b = cli.Sheet1_sf.columns.get_loc('col_b')
        self.assertEqual(cli.Sheet1_sf.iloc[0, loc_col_a].style.to_openpyxl_style(), self.sheet_1_col_a_style)
        self.assertEqual(cli.Sheet1_sf.iloc[1, loc_col_a].style.to_openpyxl_style(), self.sheet_1_col_a_cell_2_style)
        self.assertEqual(cli.Sheet1_sf.iloc[1, loc_col_b].style.to_openpyxl_style(), self.sheet_1_col_b_cell_4_style)

    def test_load_from_json_invalid_args(self):
        with self.assertRaises((TypeError, ValueError)):
            CommandLineInterface()._load_from_json()

    # noinspection PyUnresolvedReferences
    def test_init_with_json_string(self):
        with open(TEST_JSON_STRING_FILE) as f:
            json_string = f.read()
        cli = CommandLineInterface(input_json=json_string, output_path=TEST_FILENAME)
        cli.parse_as_json()
        loc_col_a = cli.Sheet1_sf.columns.get_loc('col_a')
        loc_col_b = cli.Sheet1_sf.columns.get_loc('col_b')
        self.assertEqual(cli.Sheet1_sf.iloc[0, loc_col_a].style.to_openpyxl_style(), self.sheet_1_col_a_style)
        self.assertEqual(cli.Sheet1_sf.iloc[1, loc_col_a].style.to_openpyxl_style(), self.sheet_1_col_a_cell_2_style)
        self.assertEqual(cli.Sheet1_sf.iloc[1, loc_col_b].style.to_openpyxl_style(), self.sheet_1_col_b_cell_4_style)

    @patch('sys.stderr.write')
    @patch('argparse.ArgumentParser.parse_args',
           return_value=argparse.Namespace(version=False, show_schema=False, test=False,
                                           json_path=None, json=None))
    def test_get_cli_args_invalid_args(self, args_mock, stderr_mock):
        with self.assertRaises(SystemExit):
            get_cli_args()
