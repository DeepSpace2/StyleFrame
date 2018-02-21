# coding:utf-8
from . import utils
from openpyxl.formatting import ColorScaleRule
from openpyxl.styles import PatternFill, Style, Color, Border, Side, Font, Alignment, Protection


class Styler(object):
    """
    Creates openpyxl Style to be applied
    """
    def __init__(self, bg_color=None, bold=False, font=utils.fonts.arial, font_size=12, font_color=None,
                 number_format=utils.number_formats.general, protection=False, underline=None,
                 border_type=utils.borders.thin, horizontal_alignment=utils.horizontal_alignments.center,
                 vertical_alignment=utils.vertical_alignments.center, wrap_text=True, shrink_to_fit=True,
                 fill_pattern_type=utils.fill_pattern_types.solid, indent=0):

        def get_color_from_string(color_str, default_color=None):
            if color_str and color_str.startswith('#'):
                color_str = color_str[1:]
            if not utils.is_hex_color_string(hex_string=color_str):
                color_str = utils.colors.get(color_str, default_color)
            return color_str

        self.bold = bold
        self.font = font
        self.font_size = font_size
        self.number_format = number_format
        self.protection = protection
        self.underline = underline
        self.border_type = border_type
        self.horizontal_alignment = horizontal_alignment
        self.vertical_alignment = vertical_alignment
        self.bg_color = get_color_from_string(bg_color, default_color=utils.colors.white)
        self.font_color = get_color_from_string(font_color, default_color=utils.colors.black)
        self.shrink_to_fit = shrink_to_fit
        self.wrap_text = wrap_text
        self.fill_pattern_type = fill_pattern_type
        self.indent = indent

    @classmethod
    def default_header_style(cls):
        return cls(bold=True)

    def create_style(self):
        side = Side(border_style=self.border_type, color=utils.colors.black)
        border = Border(left=side, right=side, top=side, bottom=side)
        return Style(font=Font(name=self.font, size=self.font_size, color=Color(self.font_color),
                               bold=self.bold, underline=self.underline),
                     fill=PatternFill(patternType=self.fill_pattern_type, fgColor=self.bg_color),
                     alignment=Alignment(horizontal=self.horizontal_alignment, vertical=self.vertical_alignment,
                                         wrap_text=self.wrap_text, shrink_to_fit=self.shrink_to_fit, indent=self.indent),
                     border=border,
                     number_format=self.number_format,
                     protection=Protection(locked=self.protection))


class ColorScaleConditionalFormatRule(object):
    """Creates a color scale conditional format rule. Wraps openpyxl's ColorScaleRule.
    Mostly should not be used directly, but through StyleFrame.add_color_scale_conditional_formatting
    """
    def __init__(self, start_type, start_value, start_color, end_type, end_value, end_color,
                 mid_type=None, mid_value=None, mid_color=None, columns_range=None):

        self.columns = columns_range

        # checking against None explicitly since mid_value may be 0
        if all(val is not None for val in (mid_type, mid_value, mid_color)):
            self.rule = ColorScaleRule(start_type=start_type, start_value=start_value,
                                       start_color=Color(start_color),
                                       mid_type=mid_type, mid_value=mid_value,
                                       mid_color=Color(mid_color),
                                       end_type=end_type, end_value=end_value,
                                       end_color=Color(end_color))
        else:
            self.rule = ColorScaleRule(start_type=start_type, start_value=start_value,
                                       start_color=Color(start_color),
                                       end_type=end_type, end_value=end_value,
                                       end_color=Color(end_color))
