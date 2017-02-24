import unittest
from StyleFrame import StyleFrame, Styler, utils
import pandas as pd
from functools import partial
import os


class StyleFrameTest(unittest.TestCase):
    TEST_FILENAME = 'styleframe_test.xlsx'

    @classmethod
    def setUpClass(cls):
        cls.ew = StyleFrame.ExcelWriter(StyleFrameTest.TEST_FILENAME)
        cls.style_kwargs = dict(bg_color=utils.colors.blue, bold=True, font='Impact',
                                font_color=utils.colors.yellow, font_size=20,
                                underline=utils.underline.single)
        cls.styler_obj = Styler(**cls.style_kwargs)
        cls.openpy_style_obj = cls.styler_obj.create_style()

    def setUp(self):
        self.sf = StyleFrame({'a': [1, 2, 3], 'b': [1, 2, 3]})
        self.apply_column_style = partial(self.sf.apply_column_style, **self.style_kwargs)
        self.apply_style_by_indexes = partial(self.sf.apply_style_by_indexes, **self.style_kwargs)
        self.apply_headers_style = partial(self.sf.apply_headers_style, **self.style_kwargs)

    @classmethod
    def tearDownClass(cls):
        try:
            os.remove(StyleFrameTest.TEST_FILENAME)
        except (OSError, FileNotFoundError, PermissionError) as ex:
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

    def test_init_dataframe(self):
        self.assertIsInstance(StyleFrame(pd.DataFrame({'a': [1, 2, 3], 'b': [1, 2, 3]})), StyleFrame)
        self.assertIsInstance(StyleFrame(pd.DataFrame()), StyleFrame)

    def test_init_styleframe(self):
        self.assertIsInstance(StyleFrame(StyleFrame({'a': [1, 2, 3]})), StyleFrame)

    def test_len(self):
        self.assertEqual(len(self.sf), len(self.sf.data_df))
        self.assertEqual(len(self.sf), 3)

    def test_apply_column_style(self):
        self.apply_column_style(cols_to_style=['a'])
        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj and self.sf.ix[
            index, 'b'].style != self.openpy_style_obj
                             for index in self.sf.index]))

        sheet = self.export_and_get_default_sheet()

        # range starts from 2 since we don't want to check the header's style
        self.assertTrue(all(sheet.cell(row=i, column=1).style == self.openpy_style_obj for i in range(2, len(self.sf))))

    def test_apply_style_by_indexes_single_col(self):
        self.apply_style_by_indexes(self.sf[self.sf['a'] == 2], cols_to_style=['a'])

        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj
                             for index in self.sf.index if self.sf.ix[index, 'a'] == 2]))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=1).style == self.openpy_style_obj for i in range(1, len(self.sf))
                            if sheet.cell(row=i, column=1).value == 2))

    def test_apply_style_by_indexes_all_cols(self):
        self.apply_style_by_indexes(self.sf[self.sf['a'] == 2])

        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj
                             for index in self.sf.index if self.sf.ix[index, 'a'] == 2]))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=j).style == self.openpy_style_obj
                            for i in range(1, len(self.sf))
                            for j in range(1, len(self.sf.columns))
                            if sheet.cell(row=i, column=1).value == 2))

    def test_apply_headers_style(self):
        self.apply_headers_style()
        self.assertEqual(self.sf.columns[0].style, self.openpy_style_obj)

        sheet = self.export_and_get_default_sheet()
        self.assertEqual(sheet.cell(row=1, column=1).style, self.openpy_style_obj)

    def test_apply_column_style_styler_obj(self):
        self.sf.apply_column_style(cols_to_style=['a'], styler_obj=self.styler_obj)
        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj
                             and self.sf.ix[index, 'b'].style != self.openpy_style_obj
                             for index in self.sf.index]))

        sheet = self.export_and_get_default_sheet()

        # range starts from 2 since we don't want to check the header's style
        self.assertTrue(all(sheet.cell(row=i, column=1).style == self.openpy_style_obj for i in range(2, len(self.sf))))

    def test_apply_style_by_indexes_single_col_styler_obj(self):
        self.sf.apply_style_by_indexes(self.sf[self.sf['a'] == 2], cols_to_style=['a'],
                                       styler_obj=self.styler_obj)

        self.assertTrue(all(self.sf.ix[index, 'a'].style == self.openpy_style_obj
                            for index in self.sf.index if self.sf.ix[index, 'a'] == 2))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=1).style == self.openpy_style_obj for i in range(1, len(self.sf))
                            if sheet.cell(row=i, column=1).value == 2))

    def test_apply_style_by_indexes_all_cols_styler_obj(self):
        self.sf.apply_style_by_indexes(self.sf[self.sf['a'] == 2], styler_obj=self.styler_obj)

        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj
                             for index in self.sf.index if self.sf.ix[index, 'a'] == 2]))

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.cell(row=i, column=j).style == self.openpy_style_obj
                            for i in range(1, len(self.sf))
                            for j in range(1, len(self.sf.columns))
                            if sheet.cell(row=i, column=1).value == 2))

    def test_apply_headers_style_styler_obj(self):
        self.sf.apply_headers_style(styler_obj=self.styler_obj)
        self.assertEqual(self.sf.columns[0].style, self.openpy_style_obj)

        sheet = self.export_and_get_default_sheet()
        self.assertEqual(sheet.cell(row=1, column=1).style, self.openpy_style_obj)

    def test_set_column_width(self):
        self.sf.set_column_width(columns=['a'], width=20)
        self.assertEqual(self.sf._columns_width['a'], 20)

        sheet = self.export_and_get_default_sheet()

        self.assertEqual(sheet.column_dimensions['A'].width, 20)

    def test_set_column_width_dict(self):
        width_dict = {'a': 20, 'b': 30}
        self.sf.set_column_width_dict(width_dict)
        self.assertEqual(self.sf._columns_width, width_dict)

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.column_dimensions[col.upper()].width == width_dict[col]
                            for col in width_dict))

    def test_set_row_height(self):
        self.sf.set_row_height(rows=[1], height=20)
        self.assertEqual(self.sf._rows_height[1], 20)

        sheet = self.export_and_get_default_sheet()

        self.assertEqual(sheet.row_dimensions[1].height, 20)

    def test_set_row_height_dict(self):
        height_dict = {1: 20, 2: 30}
        self.sf.set_row_height_dict(height_dict)
        self.assertEqual(self.sf._rows_height, height_dict)

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.row_dimensions[row].height == height_dict[row]
                            for row in height_dict))

    def test_rename(self):
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
        sf_from_excel = StyleFrame.read_excel(StyleFrameTest.TEST_FILENAME)
        # making sure content is the same
        self.assertTrue(all(list(self.sf[col]) == list(sf_from_excel[col]) for col in self.sf.columns))

    def test_read_excel_style(self):
        self.apply_headers_style()
        self.export_and_get_default_sheet(save=True)
        sf_from_excel = StyleFrame.read_excel(StyleFrameTest.TEST_FILENAME, read_style=True)
        # making sure content is the same
        self.assertTrue(all(list(self.sf[col]) == list(sf_from_excel[col]) for col in self.sf.columns))

        rows_in_excel = sf_from_excel.data_df.itertuples()
        rows_in_self = self.sf.data_df.itertuples()

        # making sure styles are the same
        self.assertTrue(all(excel_cell.value == self_cell.value
                        for row_in_excel, row_in_self in zip(rows_in_excel, rows_in_self)
                        for excel_cell, self_cell in zip(row_in_excel, row_in_self)))


def run():
    suite = unittest.TestLoader().loadTestsFromTestCase(StyleFrameTest)
    unittest.TextTestRunner().run(suite)


if __name__ == '__main__':
    unittest.main()
