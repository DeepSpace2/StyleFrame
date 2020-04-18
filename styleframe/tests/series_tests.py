import unittest
import pandas as pd

from pandas.testing import assert_frame_equal

from styleframe import StyleFrame, Styler, Container, Series, utils


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

    def test_style_accessor(self):
        sf = StyleFrame({'a': list(range(10))})
        sf.apply_style_by_indexes(sf[sf['a'] % 2 == 0], styler_obj=Styler(bold=True, bg_color=utils.colors.yellow),
                                  complement_style=Styler(bold=False, font=utils.fonts.calibri))

        control_sf = StyleFrame({'a': list(range(0, 10, 2))})
        test_sf = StyleFrame(sf.loc[sf['a'].style.font == utils.fonts.arial].reset_index(drop=True))
        assert_frame_equal(control_sf.data_df, test_sf.data_df)

        control_sf = StyleFrame({'a': list(range(0, 10, 2))})
        test_sf = StyleFrame(sf.loc[sf['a'].style.bg_color == utils.colors.yellow].reset_index(drop=True))
        assert_frame_equal(control_sf.data_df, test_sf.data_df)

        control_sf = StyleFrame({'a': list(range(0, 10, 2))})
        test_sf = StyleFrame(sf.loc[(sf['a'].style.bg_color == utils.colors.yellow)
                                    &
                                    sf['a'].style.font].reset_index(drop=True))
        assert_frame_equal(control_sf.data_df, test_sf.data_df)

        control_sf = StyleFrame({'a': list(range(1, 10, 2))})
        test_sf = StyleFrame(sf.loc[sf['a'].style.font == utils.fonts.calibri].reset_index(drop=True))
        assert_frame_equal(control_sf.data_df, test_sf.data_df)

        control_sf = StyleFrame({'a': list(range(1, 10, 2))})
        test_sf = StyleFrame(sf.loc[~sf['a'].style.bold].reset_index(drop=True))
        assert_frame_equal(control_sf.data_df, test_sf.data_df)

        control_sf = StyleFrame({'a': list(range(1, 10, 2))})
        test_sf = StyleFrame(sf.loc[~sf['a'].style.bold
                                    &
                                    (sf['a'].style.font == utils.fonts.calibri)].reset_index(drop=True))
        assert_frame_equal(control_sf.data_df, test_sf.data_df)
