import unittest

from StyleFrame import Styler, utils


class StylerTests(unittest.TestCase):
    def test_from_openpyxl_style(self):
        styler_obj = Styler(bg_color=utils.colors.yellow, bold=True, font=utils.fonts.david, font_size=16,
                            font_color=utils.colors.blue, number_format=utils.number_formats.date, protection=True,
                            underline=utils.underline.double, border_type=utils.borders.double,
                            horizontal_alignment=utils.horizontal_alignments.center,
                            vertical_alignment=utils.vertical_alignments.bottom, wrap_text=False, shrink_to_fit=True,
                            fill_pattern_type=utils.fill_pattern_types.gray0625, indent=1)

        self.assertEqual(styler_obj, Styler.from_openpyxl_style(styler_obj.to_openpyxl_style(), []))
