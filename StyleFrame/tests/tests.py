import unittest
import pandas as pd
import os

from functools import partial
from StyleFrame import CommandLineInterface, Container, StyleFrame, Styler, Series, utils

TEST_FILENAME = 'styleframe_test.xlsx'
TEST_JSON_FILE = 'test_json.json'
TEST_JSON_STRING_FILE = 'test_json_string.json'


class StyleFrameTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ew = StyleFrame.ExcelWriter(TEST_FILENAME)
        cls.styler_obj = Styler(bg_color=utils.colors.blue, bold=True, font='Impact', font_color=utils.colors.yellow,
                                font_size=20, underline=utils.underline.single)
        cls.openpy_style_obj = cls.styler_obj.create_style()

    def setUp(self):
        self.sf = StyleFrame({'a': [1, 2, 3], 'b': [1, 2, 3]})
        self.apply_column_style = partial(self.sf.apply_column_style, styler_obj=self.styler_obj, width=10)
        self.apply_style_by_indexes = partial(self.sf.apply_style_by_indexes, styler_obj=self.styler_obj, height=10)
        self.apply_headers_style = partial(self.sf.apply_headers_style, styler_obj=self.styler_obj)

    @classmethod
    def tearDownClass(cls):
        try:
            if os.path.exists(TEST_FILENAME):
                os.remove(TEST_FILENAME)
        except OSError as ex:
            print(ex)

    def export_and_get_default_sheet(self, save=False):
        self.sf.to_excel(excel_writer=self.ew, right_to_left=True, columns_to_hide=self.sf.columns[0],
                         row_to_add_filters=0, columns_and_rows_to_freeze='A2', allow_protection=True)
        if save:
            self.ew.save()
        return self.ew.book.get_sheet_by_name('Sheet1')

    def test_init_styler_obj(self):
        self.sf = StyleFrame({'a': [1, 2, 3], 'b': [1, 2, 3]}, styler_obj=self.styler_obj)

        self.assertTrue(all(self.sf.ix[index, 'a'].style == self.openpy_style_obj
                            for index in self.sf.index))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=j).style == self.openpy_style_obj
                            for i in range(2, len(self.sf))
                            for j in range(1, len(self.sf.columns))))

        with self.assertRaises(TypeError):
            StyleFrame({}, styler_obj=1)

    def test_init_dataframe(self):
        self.assertIsInstance(StyleFrame(pd.DataFrame({'a': [1, 2, 3], 'b': [1, 2, 3]})), StyleFrame)
        self.assertIsInstance(StyleFrame(pd.DataFrame()), StyleFrame)

    def test_init_styleframe(self):
        self.assertIsInstance(StyleFrame(StyleFrame({'a': [1, 2, 3]})), StyleFrame)

        with self.assertRaises(TypeError):
            StyleFrame({}, styler_obj=1)

    def test_len(self):
        self.assertEqual(len(self.sf), len(self.sf.data_df))
        self.assertEqual(len(self.sf), 3)

    def test_str(self):
        self.assertEqual(str(self.sf), str(self.sf.data_df))

    def test__get_item__(self):
        self.assertEqual(self.sf['a'].tolist(), self.sf.data_df['a'].tolist())
        self.assertTrue(self.sf.data_df[['a', 'b']].equals(self.sf[['a', 'b']].data_df))

    def test__getattr__(self):
        self.assertEqual(self.sf.fillna, self.sf.data_df.fillna)

        with self.assertRaises(AttributeError):
            self.sf.non_exisiting_method()

    def test_apply_column_style(self):
        # testing some edge cases
        with self.assertRaises(TypeError):
            self.sf.apply_column_style(cols_to_style='a', styler_obj=0)

        with self.assertRaises(KeyError):
            self.sf.apply_column_style(cols_to_style='non_existing_col', styler_obj=Styler())

        # actual tests
        self.apply_column_style(cols_to_style=['a'])
        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj
                             and self.sf.ix[index, 'b'].style != self.openpy_style_obj
                             for index in self.sf.index]))

        sheet = self.export_and_get_default_sheet()

        self.assertEqual(sheet.column_dimensions['A'].width, 10)

        # range starts from 2 since we don't want to check the header's style
        self.assertTrue(all(sheet.cell(row=i, column=1).style == self.openpy_style_obj for i in range(2, len(self.sf))))

    def test_apply_style_by_indexes_single_col(self):
        with self.assertRaises(TypeError):
            self.sf.apply_style_by_indexes(indexes_to_style=0, styler_obj=0)

        self.apply_style_by_indexes(self.sf[self.sf['a'] == 2], cols_to_style=['a'])

        self.assertTrue(all(self.sf.ix[index, 'a'].style == self.openpy_style_obj
                            for index in self.sf.index if self.sf.ix[index, 'a'] == 2))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=1).style == self.openpy_style_obj for i in range(1, len(self.sf))
                            if sheet.cell(row=i, column=1).value == 2))

        self.assertEqual(sheet.row_dimensions[3].height, 10)

    def test_apply_style_by_indexes_all_cols(self):
        self.apply_style_by_indexes(self.sf[self.sf['a'] == 2])

        self.assertTrue(all(self.sf.ix[index, 'a'].style == self.openpy_style_obj
                            for index in self.sf.index if self.sf.ix[index, 'a'] == 2))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=j).style == self.openpy_style_obj
                            for i in range(1, len(self.sf))
                            for j in range(1, len(self.sf.columns))
                            if sheet.cell(row=i, column=1).value == 2))

    def test_apply_style_by_indexes_with_single_index(self):
        self.apply_style_by_indexes(self.sf.index[0])

        self.assertTrue(all(self.sf.ix[0, col].style == self.openpy_style_obj
                            for col in self.sf.columns))

        sheet = self.export_and_get_default_sheet()

        # row=2 since sheet start from row 1 and the headers are row 1
        self.assertTrue(all(sheet.cell(row=2, column=col).style == self.openpy_style_obj
                            for col in range(1, len(self.sf.columns))))

    def test_apply_style_by_indexes_all_cols_with_multiple_indexes(self):
        self.apply_style_by_indexes([1, 2])

        self.assertTrue(all(self.sf.ix[index, col].style == self.openpy_style_obj
                            for index in [1, 2]
                            for col in self.sf.columns))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=j).style == self.openpy_style_obj
                            for i in [3, 4]  # sheet start from row 1 and headers are row 1
                            for j in range(1, len(self.sf.columns))))

    def test_apply_headers_style(self):
        self.apply_headers_style()
        self.assertEqual(self.sf.columns[0].style, self.openpy_style_obj)

        sheet = self.export_and_get_default_sheet()
        self.assertEqual(sheet.cell(row=1, column=1).style, self.openpy_style_obj)

    def test_set_column_width(self):
        # testing some edge cases
        with self.assertRaises(TypeError):
            self.sf.set_column_width(columns='a', width='a')
        with self.assertRaises(ValueError):
            self.sf.set_column_width(columns='a', width=-1)

        # actual tests
        self.sf.set_column_width(columns=['a'], width=20)
        self.assertEqual(self.sf._columns_width['a'], 20)

        sheet = self.export_and_get_default_sheet()

        self.assertEqual(sheet.column_dimensions['A'].width, 20)

    def test_set_column_width_dict(self):
        with self.assertRaises(TypeError):
            self.sf.set_column_width_dict(None)

        width_dict = {'a': 20, 'b': 30}
        self.sf.set_column_width_dict(width_dict)
        self.assertEqual(self.sf._columns_width, width_dict)

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.column_dimensions[col.upper()].width == width_dict[col]
                            for col in width_dict))

    def test_set_row_height(self):
        # testing some edge cases
        with self.assertRaises(TypeError):
            self.sf.set_row_height(rows=[1], height='a')
        with self.assertRaises(ValueError):
            self.sf.set_row_height(rows=[1], height=-1)
        with self.assertRaises(ValueError):
            self.sf.set_row_height(rows=['a'], height=-1)

        # actual tests
        self.sf.set_row_height(rows=[1], height=20)
        self.assertEqual(self.sf._rows_height[1], 20)

        sheet = self.export_and_get_default_sheet()

        self.assertEqual(sheet.row_dimensions[1].height, 20)

    def test_set_row_height_dict(self):
        with self.assertRaises(TypeError):
            self.sf.set_row_height_dict(None)

        height_dict = {1: 20, 2: 30}
        self.sf.set_row_height_dict(height_dict)
        self.assertEqual(self.sf._rows_height, height_dict)

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.row_dimensions[row].height == height_dict[row]
                            for row in height_dict))

    def test_rename(self):
        with self.assertRaises(TypeError):
            self.sf.rename(columns=None)

        original_columns_name = list(self.sf.columns)

        names_dict = {'a': 'A', 'b': 'B'}
        # testing rename with inplace = True
        self.sf.rename(columns=names_dict, inplace=True)

        self.assertTrue(all(new_col_name in self.sf.columns
                            for new_col_name in names_dict.values()))

        new_columns_name = list(self.sf.columns)
        # check that the columns order did not change after renaming
        self.assertTrue(all(original_columns_name.index(old_col_name) == new_columns_name.index(new_col_name)
                            for old_col_name, new_col_name in names_dict.items()))

        # using the old name should raise a KeyError
        with self.assertRaises(KeyError):
            # noinspection PyStatementEffect
            self.sf['a']

        # testing rename with inplace = False
        names_dict = {v: k for k, v in names_dict.items()}
        new_sf = self.sf.rename(columns=names_dict, inplace=False)
        self.assertTrue(all(new_col_name in new_sf.columns
                            for new_col_name in names_dict.values()))

        # using the old name should raise a KeyError
        with self.assertRaises(KeyError):
            # noinspection PyStatementEffect
            new_sf['A']

    def test_read_excel_no_style(self):
        self.apply_headers_style()
        self.export_and_get_default_sheet(save=True)
        sf_from_excel = StyleFrame.read_excel(TEST_FILENAME)
        # making sure content is the same
        self.assertTrue(all(list(self.sf[col]) == list(sf_from_excel[col]) for col in self.sf.columns))

    def test_read_excel_style(self):
        self.apply_headers_style()
        self.export_and_get_default_sheet(save=True)
        sf_from_excel = StyleFrame.read_excel(TEST_FILENAME, read_style=True)
        # making sure content is the same
        self.assertTrue(all(list(self.sf[col]) == list(sf_from_excel[col]) for col in self.sf.columns))

        rows_in_excel = sf_from_excel.data_df.itertuples()
        rows_in_self = self.sf.data_df.itertuples()

        # making sure styles are the same
        self.assertTrue(all(excel_cell.value == self_cell.value
                        for row_in_excel, row_in_self in zip(rows_in_excel, rows_in_self)
                        for excel_cell, self_cell in zip(row_in_excel, row_in_self)))

    def test_row_indexes(self):
        self.assertEqual(self.sf.row_indexes, (1, 2, 3, 4))


class ContainerTest(unittest.TestCase):
    def setUp(self):
        self.cont_1 = Container(1)
        self.cont_2 = Container(2)

    def test__gt__(self):
        self.assertGreater(self.cont_2, self.cont_1)
        self.assertGreater(self.cont_2, 1)

    def test__ge__(self):
        self.assertGreaterEqual(self.cont_1, self.cont_1)
        self.assertGreaterEqual(self.cont_2, self.cont_1)
        self.assertFalse(self.cont_1 >= self.cont_2)
        self.assertFalse(self.cont_1 >= 3)

    def test__lt__(self):
        self.assertLess(self.cont_1, self.cont_2)
        self.assertLess(self.cont_2, 3)
        self.assertFalse(self.cont_2 < self.cont_1)

    def test__le__(self):
        self.assertLessEqual(self.cont_1, self.cont_1)
        self.assertLessEqual(self.cont_1, self.cont_2)
        self.assertLessEqual(self.cont_1, 3)
        self.assertFalse(self.cont_2 < self.cont_1)

    def test__add__(self):
        self.assertEqual(self.cont_1 + self.cont_1, self.cont_2)
        self.assertEqual(self.cont_1 + 1, self.cont_2)

    def test__sub__(self):
        self.assertEqual(self.cont_2 - self.cont_1, self.cont_1)
        self.assertEqual(self.cont_2 - 1, self.cont_1)

    def test__div__(self):
        self.assertEqual(self.cont_2 / self.cont_2, self.cont_1)
        self.assertEqual(self.cont_2 / self.cont_1, self.cont_2)
        self.assertEqual(self.cont_2 / 3, Container(2/3))

    def test__mul__(self):
        self.assertEqual(self.cont_1 * self.cont_1, self.cont_1)
        self.assertEqual(self.cont_2 * 1, self.cont_2)

    def test__mod__(self):
        self.assertEqual(self.cont_2 % self.cont_1, Container(0))
        self.assertEqual(self.cont_2 % 1, Container(0))

    def test__pow__(self):
        self.assertEqual(self.cont_2 ** 2, Container(4))

    def test__int__(self):
        self.assertEqual(int(self.cont_2), 2)

    def test__float__(self):
        self.assertEqual(float(self.cont_1), 1.0)


class CommandlineInterfaceTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cli = CommandLineInterface(TEST_JSON_FILE, TEST_FILENAME)
        cls.sheet_1_col_a_style = Styler(bg_color=utils.colors.blue, font_color=utils.colors.yellow).create_style()
        cls.sheet_1_col_a_cell_2_style = Styler(bold=True, font=utils.fonts.arial, font_size=30, font_color=utils.colors.green,
                                                border_type=utils.borders.double).create_style()
        cls.sheet_1_col_b_cell_4_style = Styler(bold=True, font=utils.fonts.arial, font_size=16).create_style()

    def tearDown(self):
        try:
            os.remove(TEST_FILENAME)
        except OSError as ex:
            print(ex)

    def test_parse_as_json(self):
        self.cli.parse_as_json()
        self.assertEqual(self.cli.Sheet1_sf.ix[0, 'col_a'].style, self.sheet_1_col_a_style)
        self.assertEqual(self.cli.Sheet1_sf.ix[1, 'col_a'].style, self.sheet_1_col_a_cell_2_style)
        self.assertEqual(self.cli.Sheet1_sf.ix[1, 'col_b'].style, self.sheet_1_col_b_cell_4_style)

    def test_init_with_json_string(self):
        with open(TEST_JSON_STRING_FILE) as f:
            json_string = f.read()
        cli = CommandLineInterface(input_json=json_string, output_path=TEST_FILENAME)
        cli.parse_as_json()
        self.assertEqual(cli.Sheet1_sf.ix[0, 'col_a'].style, self.sheet_1_col_a_style)
        self.assertEqual(cli.Sheet1_sf.ix[1, 'col_a'].style, self.sheet_1_col_a_cell_2_style)
        self.assertEqual(cli.Sheet1_sf.ix[1, 'col_b'].style, self.sheet_1_col_b_cell_4_style)


class SeriesTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pandas_series = pd.Series((None, 1))
        cls.sf_series = Series((Container(None), Container(1)))

    def test_isnull(self):
        self.assertTrue(all(p_val == sf_val
                            for p_val, sf_val in zip(self.pandas_series.isnull(), self.sf_series.isnull())))

    def test_notnull(self):
        self.assertTrue(all(p_val == sf_val
                            for p_val, sf_val in zip(self.pandas_series.notnull(), self.sf_series.notnull())))


def run():
    test_classes = [ContainerTest, StyleFrameTest, CommandlineInterfaceTest, SeriesTest]
    for test_class in test_classes:
        suite = unittest.TestLoader().loadTestsFromTestCase(test_class)
        unittest.TextTestRunner().run(suite)


if __name__ == '__main__':
    unittest.main()
