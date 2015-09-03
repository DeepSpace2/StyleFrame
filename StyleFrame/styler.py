# coding:utf-8
from openpyxl.styles import PatternFill, Style, Color, Border, Side, Font, Alignment, colors as op_colors


class Styler(object):
    """
    Creates openpyxl Style to be applied
    """
    def __init__(self, bg_color='white', bold=False, font_size=12, font_color='black', number_format='General',
                 underline=None):
        self.bg_color = bg_color
        self.bold = bold
        self.font_size = font_size
        self.font_color = font_color
        self.number_format = number_format
        self.underline = underline

    def create_style(self):
        side = Side(border_style='thin', color=colors.black)
        border = Border(left=side, right=side, top=side, bottom=side)
        return Style(font=Font(name="Arial", size=self.font_size, color=Color(colors.get(self.font_color, colors.black)),
                               bold=self.bold, underline=self.underline),
                     fill=PatternFill(patternType='solid', fgColor=Color(colors.get(self.bg_color, colors.white))),
                     alignment=Alignment(horizontal='center', vertical='center', wrap_text=True, shrink_to_fit=True, indent=0),
                     border=border,
                     number_format=self.number_format)


def not_supported(*args, **kwargs):
        raise NotImplementedError('ImmutableDict is immutable')


class ImmutableDict(dict):
    __delitem__ = not_supported
    __setitem__ = not_supported
    __setattr__ = not_supported
    update = not_supported
    clear = not_supported
    pop = not_supported
    popitem = not_supported

    def __getattr__(self, item):
        return self[item]

number_formats = ImmutableDict(general='General', date='DD/MM/YY', percent='0.0%', time_24_hours='HH:MM',
                               time_12_hours='h:MM AM/PM', date_time='DD/MM/YY HH:MM',
                               thousands_comma_sep='#,##0')

colors = ImmutableDict(white='FFFFFF', blue=op_colors.BLUE, yellow=op_colors.YELLOW, green=op_colors.GREEN,
                       black=op_colors.BLACK, red=op_colors.RED, purple='800080')
