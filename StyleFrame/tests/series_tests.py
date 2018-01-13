import unittest
import pandas as pd

from StyleFrame import Container, Series


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
