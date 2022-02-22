import unittest

from styleframe import Styler, utils


class StylerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.yellow_1 = Styler(bg_color='yellow')
        cls.yellow_2 = Styler(bg_color='yellow')
        cls.blue = Styler(bg_color='blue')
        cls.bold = Styler(bold=True)
        cls.underline = Styler(underline='single')
        cls.yellow_bold_underline = Styler(bg_color='yellow', bold=True, underline='single')

    def test_eq(self):
        self.assertEqual(self.yellow_1, self.yellow_2)
        self.assertNotEqual(self.yellow_1, self.blue)

    def test_add(self):
        self.assertEqual(self.yellow_1 + self.bold + self.underline, self.yellow_bold_underline)
        self.assertEqual(self.yellow_2.bold, False)
        self.assertEqual(self.yellow_2 + Styler(), self.yellow_1)

    def test_combine(self):
        self.assertEqual(Styler.combine(self.yellow_1, self.bold, self.underline), self.yellow_bold_underline)

    def test_from_openpyxl_style(self):
        styler_obj = Styler(bg_color=utils.colors.yellow, bold=True, font=utils.fonts.david, font_size=16,
                            font_color=utils.colors.blue, number_format=utils.number_formats.date, protection=True,
                            underline=utils.underline.double, border_type=utils.borders.double,
                            horizontal_alignment=utils.horizontal_alignments.center,
                            vertical_alignment=utils.vertical_alignments.bottom, wrap_text=False, shrink_to_fit=True,
                            fill_pattern_type=utils.fill_pattern_types.gray0625, indent=1)

        self.assertEqual(styler_obj, Styler.from_openpyxl_style(styler_obj.to_openpyxl_style(), []))

    def test_default_grid_invalid_args(self):
        with self.assertRaises(ValueError):
            Styler(border_type=utils.borders.default_grid, bg_color=utils.colors.yellow)
        with self.assertRaises(ValueError):
            Styler(border_type=utils.borders.default_grid, fill_pattern_type=utils.fill_pattern_types.light_grid)
        with self.assertRaises(ValueError):
            Styler(border_type=utils.borders.default_grid, bg_color=utils.colors.yellow,
                   fill_pattern_type=utils.fill_pattern_types.light_grid)

