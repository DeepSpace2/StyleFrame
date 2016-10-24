import unittest
from StyleFrame import StyleFrame, Styler, utils
from functools import partial


class StyleFrameTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.ew = StyleFrame.ExcelWriter('test.xlsx')
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

    def export_and_get_default_sheet(self):
        self.sf.to_excel(excel_writer=self.ew)
        return self.ew.book.get_sheet_by_name('Sheet1')

    def test_apply_column_style(self):
        self.apply_column_style(cols_to_style=['a'])
        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj and self.sf.ix[index, 'b'].style != self.openpy_style_obj
                             for index in self.sf.index]))

        sheet = self.export_and_get_default_sheet()

        # range starts from 1 since we don't want to check the header's style
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

        # range starts from 1 since we don't want to check the header's style
        self.assertTrue(all(sheet.cell(row=i, column=1).style == self.openpy_style_obj for i in range(2, len(self.sf))))

    def test_apply_style_by_indexes_single_col_styler_obj(self):
        self.sf.apply_style_by_indexes(self.sf[self.sf['a'] == 2], cols_to_style=['a'],
                                       styler_obj=self.styler_obj)

        self.assertTrue(all([self.sf.ix[index, 'a'].style == self.openpy_style_obj
                             for index in self.sf.index if self.sf.ix[index, 'a'] == 2]))

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
        self.assertEqual(self.sf.columns_width['a'], 20)

        sheet = self.export_and_get_default_sheet()

        self.assertEqual(sheet.column_dimensions['A'].width, 20)

    def test_set_column_width_dict(self):
        width_dict = {'a': 20, 'b': 30}
        self.sf.set_column_width_dict(width_dict)
        self.assertEqual(self.sf.columns_width, width_dict)

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.column_dimensions[col.upper()].width == width_dict[col]
                            for col in width_dict))

    def test_set_row_height(self):
        self.sf.set_row_height(rows=[1], height=20)
        self.assertEqual(self.sf.rows_height[1], 20)

        sheet = self.export_and_get_default_sheet()

        self.assertEqual(sheet.row_dimensions[1].height, 20)

    def test_set_row_height_dict(self):
        height_dict = {1: 20, 2: 30}
        self.sf.set_row_height_dict(height_dict)
        self.assertEqual(self.sf.rows_height, height_dict)

        sheet = self.export_and_get_default_sheet()

        self.assertTrue(all(sheet.row_dimensions[row].height == height_dict[row]
                            for row in height_dict))

    def test_rename(self):
        names_dict = {'a': 'A', 'b': 'B'}
        self.sf.rename(columns=names_dict, inplace=True)
        self.assertTrue(all(new_col_name in self.sf.columns
                            for new_col_name in names_dict.values()))

        # using the old name should raise a KeyError
        with self.assertRaises(KeyError):
            # noinspection PyStatementEffect
            self.sf['a']


def run():
    suite = unittest.TestLoader().loadTestsFromTestCase(StyleFrameTest)
    unittest.TextTestRunner().run(suite)

if __name__ == '__main__':
    unittest.main()
