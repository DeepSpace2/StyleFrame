import unittest
from StyleFrame import StyleFrame
from StyleFrame import Styler


class StyleFrameTest(unittest.TestCase):
    def setUp(self):
        self.df = StyleFrame({'a': [1, 2, 3], 'b': [1, 2, 3]})

    def apply_column_style_test(self):
        self.df.apply_column_style(cols_to_style=['a'], bg_color='blue', bold=True)
        self.assertTrue(all([self.df.ix[index, 'a'].style == Styler(bg_color='blue', bold=True).create_style()
                             and self.df.ix[index, 'b'].style != Styler(bg_color='blue', bold=True).create_style()
                             for index in self.df.index]))

    def apply_style_by_indexes_test(self):
        self.df.apply_style_by_indexes(self.df[self.df['a'] == 2], cols_to_style=['a'], bg_color='blue')
        self.assertTrue(all([self.df.ix[index, 'a'].style == Styler(bg_color='blue').create_style() for index in self.df.index if self.df.ix[index, 'a'] == 2]))
